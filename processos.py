# -*- coding: utf-8 -*-
import os
import time
import tempfile
import threading
import concurrent.futures
from pathlib import Path
from typing import Optional

import fitz
import gradio as gr
from google import genai
from google.genai import types

from report_template_processos import REPORT_TEMPLATE_INSTRUCTIONS, SYSTEM_PROMPT as SYSTEM_PROMPT_PROC
from utils import _retry, _gerar_docx, _responder_pergunta_generica, _get_clients_proc, _barra_progresso, _comprimir_pdf

CHUNK_MAX_PAGES_PROC    = 400
MODO_DIRETO_MAX_PROC    = 700
MODO_DIRETO_MAX_MB_PROC = 45
COMPRESSAO_PRE_MB_PROC  = 50   # só pré-comprime se puder caber no modo direto
MODEL_PROC_EXTR = os.getenv("GEMINI_MODEL_EXTRACAO", "gemini-2.5-flash")
MODEL_PROC_FAST = os.getenv("GEMINI_MODEL_RAPIDO", "gemini-2.5-flash")
MODEL_PROC_PRO = os.getenv("GEMINI_MODEL_CONSOLIDACAO", "gemini-2.5-pro")
MODEL_PROC_CONS = MODEL_PROC_PRO
# ══════════════════════════════════════════════════════════════════════════════
# MÓDULO — ANÁLISE DE PROCESSOS (com melhorias Teste B)
# ══════════════════════════════════════════════════════════════════════════════

_proc_cache_nome: dict[str, str] = {}
_proc_cache_lock = threading.Lock()

PROMPT_EXTR_PROC = (
    "Voce esta analisando uma parte de um processo judicial brasileiro.\n"
    "IMPORTANTE: este arquivo pode ser um fragmento de um processo maior dividido em multiplos PDFs "
    "pelo tribunal, OU pode ser um processo independente. Extraia o numero do processo se disponivel — "
    "isso sera usado para detectar continuidade na consolidacao. Trate cada fragmento como parte de "
    "um documento continuo; a consolidacao final verificara se sao o mesmo processo ou processos distintos.\n"
    "Extraia e registre COM MAXIMO DETALHE todas as informacoes presentes nestas paginas:\n"
    "- Partes (nomes, CPF/CNPJ, qualidade: exequente, executado, avalista, etc.)\n"
    "- Advogados e OABs\n"
    "- Datas de todos os atos\n"
    "- Valores monetarios (exatamente como constam)\n"
    "- Indices contratuais: juros remuneratorios, moratorios, correcao monetaria, multa\n"
    "- Identificacao dos titulos: numero da CCB, contrato, duplicata, etc.\n"
    "- Garantias e seus detalhes\n"
    "- Assinaturas e representantes\n"
    "- Citacoes, intimacoes e suas modalidades\n"
    "- Excecoes, embargos e recursos: teses e status\n"
    "- Decisoes e despachos\n"
    "- Todos os andamentos processuais com data e referencia (Mov./fls.)\n\n"
    "NAO formate como relatorio final ainda. Apenas extraia tudo com fidelidade."
)


def _proc_dividir_pdf(pdf_path: str) -> list:
    try:
        doc = fitz.open(pdf_path)
        total = len(doc)
        if total <= CHUNK_MAX_PAGES_PROC:
            doc.close()
            return [(pdf_path, 0, total)]
        chunks = []
        for start in range(0, total, CHUNK_MAX_PAGES_PROC):
            end = min(start + CHUNK_MAX_PAGES_PROC, total)
            sub = fitz.open()
            sub.insert_pdf(doc, from_page=start, to_page=end - 1)
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            sub.save(tmp.name, garbage=4, deflate=True)
            sub.close()
            tmp.close()
            chunks.append((tmp.name, start, total))
        doc.close()
        return chunks
    except Exception as e:
        return [(pdf_path, 0, 0)]


def _proc_extrair_chunk(args) -> tuple:
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    pg_fim = min(offset + CHUNK_MAX_PAGES_PROC, total_pg) if total_pg else "?"

    textos_pag = []
    imagens_pag = []
    total_img_bytes = 0
    MAX_IMG_BYTES = 19 * 1024 * 1024

    try:
        doc = fitz.open(chunk_path)
        for i, page in enumerate(doc):
            txt = page.get_text().strip()
            pg = offset + i + 1
            alfa = sum(c.isalpha() for c in txt)
            if len(txt) >= 50 and alfa / len(txt) >= 0.4 and len(txt.split()) >= 30:
                textos_pag.append(f"[p.{pg}]\n{txt}")
            else:
                if total_img_bytes < MAX_IMG_BYTES:
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
                    img = pix.tobytes("jpeg", jpg_quality=50)
                    total_img_bytes += len(img)
                    imagens_pag.append((pg, img))
        doc.close()
    except Exception:
        pass

    n_scan = len(imagens_pag)

    cabecalho = (
        f"[PARTE {idx+1}/{n_total} — paginas {pg_ini}-{pg_fim}]\n"
        "Nao omita nenhum dado mesmo que pareca repetitivo.\n"
    )
    prompt_txt = cabecalho + PROMPT_EXTR_PROC
    if textos_pag:
        prompt_txt += "\n\nTEXTO PESQUISAVEL EXTRAIDO:\n" + "\n\n".join(textos_pag)
    if n_scan:
        prompt_txt += f"\n\n{n_scan} pagina(s) escaneada(s) enviadas como imagem — aplique OCR visual completo."

    all_parts: list = [types.Part(text=prompt_txt)]
    for pg_num, img_bytes in imagens_pag:
        all_parts.append(types.Part(text=f"[Pagina {pg_num} — escaneada]"))
        all_parts.append(types.Part(inline_data=types.Blob(mime_type="image/jpeg", data=img_bytes)))

    def _call():
        contents = [types.Content(role="user", parts=all_parts)]
        cfg = types.GenerateContentConfig(
            thinking_config=types.ThinkingConfig(thinking_budget=0)
        )
        resp = client.models.generate_content(model=MODEL_PROC_EXTR, contents=contents, config=cfg)
        return resp.text

    resultado = _retry(_call)
    nota = f"{n_scan} pag. escaneadas via imagem" if n_scan else ""
    return idx, resultado, nota


def _proc_extrair_chunk_fileapi(args) -> tuple:
    """Comprime chunk, faz upload para o File API e extrai via leitura nativa de PDF."""
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    pg_fim = min(offset + CHUNK_MAX_PAGES_PROC, total_pg) if total_pg else "?"

    # Comprimir chunk antes do upload (acontece em paralelo nos workers)
    comp_chunk, orig_mb, comp_mb = _comprimir_pdf(chunk_path)
    source_for_upload = comp_chunk
    if comp_chunk != chunk_path:
        comp_nota = f"{orig_mb:.0f}→{comp_mb:.0f}MB"
    else:
        comp_nota = f"{orig_mb:.0f}MB"

    ascii_tmp = None
    try:
        source_for_upload.encode("ascii")
        upload_path = source_for_upload
    except (UnicodeEncodeError, AttributeError):
        import shutil as _shutil
        t = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        t.close()
        _shutil.copy2(source_for_upload, t.name)
        upload_path = ascii_tmp = t.name

    arq = client.files.upload(file=upload_path)
    total_wait, wait_time = 0, 1
    while total_wait < 120:
        state_name = getattr(getattr(arq, "state", None), "name", "")
        if state_name in ("ACTIVE", "FAILED"):
            break
        time.sleep(wait_time)
        total_wait += wait_time
        arq = client.files.get(name=arq.name)
        wait_time = min(wait_time + 1, 4)

    if ascii_tmp:
        try: os.remove(ascii_tmp)
        except: pass
    if comp_chunk != chunk_path:
        try: os.remove(comp_chunk)
        except: pass

    prompt = (
        f"[PARTE {idx+1}/{n_total} — paginas {pg_ini}-{pg_fim}]\n"
        "Nao omita nenhum dado mesmo que pareca repetitivo.\n\n"
        + PROMPT_EXTR_PROC
    )
    mime = getattr(arq, "mime_type", None) or "application/pdf"
    contents = [types.Content(role="user", parts=[
        types.Part(text=prompt),
        types.Part(file_data=types.FileData(file_uri=arq.uri, mime_type=mime)),
    ])]

    _cfg_extr = types.GenerateContentConfig(
        thinking_config=types.ThinkingConfig(thinking_budget=0)
    )

    def _call():
        return client.models.generate_content(model=MODEL_PROC_EXTR, contents=contents, config=_cfg_extr).text

    resultado = _retry(_call)

    try: client.files.delete(name=arq.name)
    except: pass

    return idx, resultado, f"File API | {comp_nota}"


def _proc_obter_cache(client, model_cons: str) -> Optional[str]:
    global _proc_cache_nome
    with _proc_cache_lock:
        if model_cons in _proc_cache_nome:
            # Valida se o cache ainda existe na API (TTL de 1h pode ter expirado)
            try:
                client.caches.get(name=_proc_cache_nome[model_cons])
                return _proc_cache_nome[model_cons]
            except Exception:
                _proc_cache_nome.pop(model_cons, None)
        try:
            cache = client.caches.create(
                model=model_cons,
                config=types.CreateCachedContentConfig(
                    display_name="invista_proc_template_v1",
                    system_instruction=SYSTEM_PROMPT_PROC,
                    contents=[types.Content(
                        role="user",
                        parts=[types.Part(text=f"TEMPLATE DO RELATORIO:\n\n{REPORT_TEMPLATE_INSTRUCTIONS}")]
                    )],
                    ttl="3600s",
                )
            )
            _proc_cache_nome[model_cons] = cache.name
            return cache.name
        except Exception:
            return None


def _proc_consolidar(client, parciais: list, instrucoes: str, cache_name, model_cons: str, versao_resumida: bool = False) -> str:
    blocos = "\n\n".join(
        f"{'='*60}\nPARTE {i+1}/{len(parciais)}\n{'='*60}\n{p}"
        for i, p in enumerate(parciais)
    )

    if versao_resumida:
        # Skip cache (cache contains full template), use only resumo template
        instrucoes_extras = f"\n\nINSTRUCOES ADICIONAIS: {instrucoes.strip()}" if instrucoes.strip() else ""
        prompt_full = (
            SYSTEM_PROMPT_PROC + "\n\n"
            + "Você está em MODO RESUMO. Ignore qualquer template detalhado e siga ESTRITAMENTE o formato abaixo.\n\n"
            + TEMPLATE_RESUMIDO_PROC
            + instrucoes_extras
            + f"\n\nA seguir estão {len(parciais)} PARTE(S) de extrações brutas do processo.\n"
            + "INSTRUÇÃO OBRIGATÓRIA DE LEITURA:\n"
            + "1. Leia CADA PARTE individualmente, do começo ao fim, sem pular nenhuma.\n"
            + "2. Extraia todos os andamentos de CADA UMA das partes (PARTE 1, PARTE 2, ..., PARTE N).\n"
            + "3. Ao final, mescle todos os andamentos em UMA única lista cronológica global.\n"
            + "4. Omitir andamentos de qualquer PARTE é ERRO GRAVE — antes de finalizar, confira se a quantidade de andamentos extraídos é compatível com a totalidade das partes.\n\n"
            + blocos
        )
        def _fn():
            return client.models.generate_content(
                model=model_cons, contents=[prompt_full]
            ).text
        return _retry(_fn)

    _instrucoes_extra = f"INSTRUCOES ADICIONAIS: {instrucoes.strip()}\n\n" if instrucoes.strip() else ""
    _aviso_multiplos = (
        f"ATENCAO: foram carregados {len(parciais)} fragmentos de arquivo(s). "
        "Verifique se compartilham o mesmo numero de processo — se sim, trate como UM UNICO PROCESSO CONTINUO. "
        "Se forem processos distintos, gere secoes separadas para cada um.\n\n"
    ) if len(parciais) > 1 else ""
    prompt = (
        "A seguir estao as extracoes brutas de cada parte do processo. "
        "Com base nelas, produza o relatorio juridico final seguindo rigorosamente "
        "o modelo de formatacao.\n\n"
        + _aviso_multiplos
        + _instrucoes_extra
        + blocos
    )

    if cache_name:
        contents = [types.Content(role="user", parts=[types.Part(text=prompt)])]
        config = types.GenerateContentConfig(cached_content=cache_name)
        def _fn():
            return client.models.generate_content(
                model=model_cons, contents=contents, config=config
            ).text
        try:
            return _retry(_fn)
        except Exception as e:
            msg = str(e)
            if "PERMISSION_DENIED" in msg or "CachedContent not found" in msg or "403" in msg:
                # Cache invalido — descarta e refaz sem cache
                with _proc_cache_lock:
                    _proc_cache_nome.pop(model_cons, None)
                prompt_full = SYSTEM_PROMPT_PROC + "\n\n" + REPORT_TEMPLATE_INSTRUCTIONS + "\n\n" + prompt
                def _fn_fallback():
                    return client.models.generate_content(
                        model=model_cons, contents=[prompt_full]
                    ).text
                return _retry(_fn_fallback)
            raise
    else:
        prompt_full = SYSTEM_PROMPT_PROC + "\n\n" + REPORT_TEMPLATE_INSTRUCTIONS + "\n\n" + prompt
        def _fn():
            return client.models.generate_content(
                model=model_cons, contents=[prompt_full]
            ).text
        return _retry(_fn)


def _proc_processar_relacionados(pdf_paths: list, client1, client2, instrucoes: str, cache_name, model_cons: str) -> str:
    """Extrai e consolida processos relacionados, gerando seções B.1, B.2, etc."""
    try:
        texto_relacionados = ""
        for idx, path in enumerate(pdf_paths):
            chunks = _proc_dividir_pdf(path)
            cli = client1 if idx % 2 == 0 else client2
            for chunk_path, offset, total in chunks:
                try:
                    _, res, _ = _proc_extrair_chunk((idx, chunk_path, offset, total, len(chunks), cli))
                    if res:
                        texto_relacionados += f"\n--- {Path(path).name} ---\n{res}"
                except Exception:
                    pass

        if not texto_relacionados.strip():
            return ""

        prompt = (
            "Com base nos processos relacionados abaixo, "
            "gere as secoes de processos relacionados (B.1, B.2, etc.) "
            "seguindo rigorosamente o template.\n\n"
            "REGRAS:\n"
            "- Recursos e Embargos a Execucao: integre na analise do processo principal\n"
            "- IDPJs, Paulianas, Embargos de Terceiro: crie secoes B proprias\n\n"
            f"PROCESSOS RELACIONADOS:\n{texto_relacionados}\n\n"
        )
        if instrucoes.strip():
            prompt += f"INSTRUCOES: {instrucoes.strip()}\n\n"
        prompt += "Gere APENAS as secoes de processos relacionados."

        if cache_name:
            contents = [types.Content(role="user", parts=[types.Part(text=prompt)])]
            config = types.GenerateContentConfig(cached_content=cache_name)
            def _fn():
                return client1.models.generate_content(
                    model=model_cons, contents=contents, config=config
                ).text
            try:
                return _retry(_fn)
            except Exception as e:
                msg = str(e)
                if "PERMISSION_DENIED" in msg or "CachedContent not found" in msg or "403" in msg:
                    with _proc_cache_lock:
                        _proc_cache_nome.pop(model_cons, None)
                    prompt_full = SYSTEM_PROMPT_PROC + "\n\n" + REPORT_TEMPLATE_INSTRUCTIONS + "\n\n" + prompt
                    def _fn_fallback():
                        return client1.models.generate_content(
                            model=model_cons, contents=[prompt_full]
                        ).text
                    return _retry(_fn_fallback)
                raise
        else:
            prompt_full = SYSTEM_PROMPT_PROC + "\n\n" + REPORT_TEMPLATE_INSTRUCTIONS + "\n\n" + prompt
            def _fn():
                return client1.models.generate_content(
                    model=model_cons, contents=[prompt_full]
                ).text
            return _retry(_fn)
    except Exception:
        return ""


TEMPLATE_RESUMIDO_PROC = """\
MODO RESUMO PROCESSUAL — siga rigorosamente este formato curto:

A. [Tipo da ação] nº [número completo] - [Vara] - [Tribunal]

• Partes:
   ∘ Exequente: [Nome] — [CPF/CNPJ]
   ∘ Executado(s): [Nome] — [CPF/CNPJ]

• Data de distribuição: DD/MM/AAAA

• Valor da causa: R$ X.XXX.XXX,XX

• Principais andamentos: (lista cronológica completa do mais antigo ao mais recente, com data, descrição objetiva do ato processual e referência Mov./fls.)
   ▪ DD/MM/AAAA — [Descrição objetiva] (Mov. X | fls. XX/XX)
   ▪ DD/MM/AAAA — [Próximo ato] (Mov. X | fls. XX/XX)
   ...

REGRAS RÍGIDAS:
- NÃO gere análise de risco, recomendações, garantias, honorários, lastro, índices contratuais, SAT, assinaturas, constrições, exceções de pré-executividade, embargos, recursos ou qualquer outra seção.
- NÃO inclua nenhum conteúdo além das 5 seções acima.
- Use os marcadores • ∘ ▪ conforme o exemplo.
- Mantenha objetividade nos andamentos.
"""


def _proc_analisar_direto(client, arquivos_gemini: list, instrucoes: str, model_cons: str, versao_resumida: bool = False) -> str:
    instrucoes_extras = ""
    if instrucoes and instrucoes.strip():
        instrucoes_extras = (
            "\n\nINSTRUCOES ADICIONAIS:\n" + instrucoes.strip()
        )
    if versao_resumida:
        partes_prompt = [
            SYSTEM_PROMPT_PROC,
            "Você está em MODO RESUMO. Ignore qualquer template detalhado e siga ESTRITAMENTE o formato abaixo.",
            TEMPLATE_RESUMIDO_PROC + instrucoes_extras,
        ]
    else:
        partes_prompt = [
            SYSTEM_PROMPT_PROC,
            "Analise integralmente todos os PDFs e produza o relatorio juridico final "
            "seguindo rigorosamente o modelo de formatacao abaixo.",
            REPORT_TEMPLATE_INSTRUCTIONS + instrucoes_extras,
        ]
    # Texto e arquivos juntos num único Content para evitar 400 INVALID_ARGUMENT
    parts = [types.Part(text="\n\n".join(partes_prompt))]
    for arq in arquivos_gemini:
        mime = getattr(arq, "mime_type", None) or "application/pdf"
        parts.append(types.Part(file_data=types.FileData(file_uri=arq.uri, mime_type=mime)))
    contents = [types.Content(role="user", parts=parts)]
    def _fn():
        return client.models.generate_content(model=model_cons, contents=contents).text
    return _retry(_fn)


def proc_analisar(pdf_files, pdf_relacionados, instrucoes: str, usar_gemini_pro: bool = False, versao_resumida: bool = False):
    if not pdf_files:
        yield "Nenhum arquivo enviado.", "", ""
        return

    try:
        client1, client2 = _get_clients_proc()
    except Exception as e:
        yield f"Erro de configuracao: {e}", "", ""
        return

    def _nat_key(p):
        return [int(x) if x.isdigit() else x.lower()
                for x in re.split(r"(\d+)", Path(p).name)]

    pdf_paths = sorted(
        [f.name if hasattr(f, "name") else str(f) for f in pdf_files],
        key=_nat_key
    )
    log = []
    relatorio_state = ""
    model_cons = MODEL_PROC_PRO if usar_gemini_pro else MODEL_PROC_FAST

    log.append(f"Arquivo(s) recebido(s): {len(pdf_paths)}")
    for p in pdf_paths:
        mb = Path(p).stat().st_size / 1_048_576
        log.append(f"   · {Path(p).name} ({mb:.1f} MB)")
    if len(pdf_paths) > 1:
        log.append(f"   → {len(pdf_paths)} arquivos — o modelo verificara se sao fragmentos do mesmo processo ou processos distintos")
    if instrucoes.strip():
        log.append("Instrucoes adicionais recebidas.")
    log.append(f"Modelo de consolidacao: {model_cons}")
    yield "\n".join(log), "", ""

    # Inspecionar e dividir
    log.append("\nInspecionando e dividindo PDFs...")
    yield "\n".join(log), "", ""

    todos_chunks = []
    _temp_comprimidos: list = []
    for path in pdf_paths:
        nome = Path(path).name
        mb_orig = Path(path).stat().st_size / 1_048_576

        path_proc = path
        if 10 <= mb_orig <= COMPRESSAO_PRE_MB_PROC:
            # Arquivo pequeno: pré-comprime pois pode caber no modo direto
            log.append(f"   Comprimindo {nome} ({mb_orig:.0f} MB)...")
            yield "\n".join(log), "", ""
            comp_path, orig_mb, comp_mb = _comprimir_pdf(path)
            if comp_path != path:
                reducao = (1 - comp_mb / orig_mb) * 100
                log[-1] = f"   {nome}: {orig_mb:.0f} MB → {comp_mb:.0f} MB (-{reducao:.0f}%)"
                yield "\n".join(log), "", ""
                _temp_comprimidos.append(comp_path)
                path_proc = comp_path
        elif mb_orig > COMPRESSAO_PRE_MB_PROC:
            # Arquivo grande: vai direto ao modo chunked; compressao ocorre por chunk em paralelo
            log.append(f"   {nome} ({mb_orig:.0f} MB) — compressao por chunk (paralelo)")
            yield "\n".join(log), "", ""

        doc_tmp = fitz.open(path_proc)
        total_pg = len(doc_tmp)
        esc = []
        for i, p in enumerate(doc_tmp):
            txt = p.get_text().strip()
            if not txt or len(txt) < 50:
                esc.append(i)
                continue
            alfa = sum(c.isalpha() for c in txt)
            if alfa / len(txt) < 0.4 or len(txt.split()) < 30:
                esc.append(i)
        doc_tmp.close()
        pct = int(len(esc) / total_pg * 100) if total_pg else 0
        info_scan = f"{len(esc)} pag. escaneadas ({pct}%)" if esc else "todas pesquisaveis"
        chunks = _proc_dividir_pdf(path_proc)
        log.append(f"   · {nome}: {total_pg} pag. — {info_scan}")
        if len(chunks) > 1:
            log.append(f"     Dividido em {len(chunks)} partes de {CHUNK_MAX_PAGES_PROC} pag.")
        for chunk_path, offset, total in chunks:
            todos_chunks.append((chunk_path, offset, total, path))

    n = len(todos_chunks)
    total_paginas = sum(c[2] for c in todos_chunks if c[2])
    compressed_total_mb = sum(Path(c[0]).stat().st_size for c in todos_chunks) / 1_048_576
    yield "\n".join(log), "", ""

    t_inicio = time.time()
    cache = None

    try:
        # Modo direto (1 chamada) para documentos pequenos
        if total_paginas <= MODO_DIRETO_MAX_PROC and compressed_total_mb <= MODO_DIRETO_MAX_MB_PROC:
            log.append(f"\nModo direto ({total_paginas} pag. · {compressed_total_mb:.0f} MB — 1 chamada ao {model_cons})")
            yield "\n".join(log), "", ""

            arquivos_gemini = []
            for idx, (chunk_path, offset, total_pg, pdf_orig) in enumerate(todos_chunks, 1):
                tam = Path(chunk_path).stat().st_size / 1_048_576
                log.append(f"   Enviando {idx}/{len(todos_chunks)} ({tam:.1f} MB)...")
                yield "\n".join(log), "", ""

                try:
                    chunk_path.encode("ascii")
                    upload_path = chunk_path
                except UnicodeEncodeError:
                    import shutil as _shutil
                    tmp_ascii = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                    tmp_ascii.close()
                    _shutil.copy2(chunk_path, tmp_ascii.name)
                    upload_path = tmp_ascii.name

                arq = client1.files.upload(file=upload_path)
                # Polling com backoff: 1s→2s→3s→4s→4s..., teto de 120s no total
                total_wait = 0
                wait_time = 1
                while total_wait < 120:
                    state = getattr(arq, "state", None)
                    state_name = getattr(state, "name", str(state) if state is not None else "")
                    if state_name in ("ACTIVE", "FAILED"):
                        break
                    time.sleep(wait_time)
                    total_wait += wait_time
                    arq = client1.files.get(name=arq.name)
                    wait_time = min(wait_time + 1, 4)
                arquivos_gemini.append(arq)
                log[-1] = f"   Arquivo {idx}/{len(todos_chunks)} pronto."
                yield "\n".join(log), "", ""

            log.append("\nGerando relatorio...")
            yield "\n".join(log), "", ""
            relatorio = _proc_analisar_direto(client1, arquivos_gemini, instrucoes, model_cons, versao_resumida)

        else:
            # Modo chunked via File API — 4 workers, leitura nativa de PDF
            log.append(f"\nModo chunked ({n} chunk(s) · {compressed_total_mb:.0f} MB · 4 workers · File API)...")
            log.append(_barra_progresso(0, n))
            yield "\n".join(log), "", ""

            parciais: dict = {}

            def _worker_proc(idx):
                cp, offset, total_pg, _ = todos_chunks[idx]
                cli = client1 if idx % 2 == 0 else client2
                return _proc_extrair_chunk_fileapi((idx, cp, offset, total_pg, n, cli))

            with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
                futures = {ex.submit(_worker_proc, i): i for i in range(n)}
                for future in concurrent.futures.as_completed(futures):
                    try:
                        idx, res, nota = future.result()
                        parciais[idx] = res
                        c = len(parciais)
                        log[-1] = _barra_progresso(c, n)
                        log.append(
                            f"   Chunk {idx+1}/{n} extraido"
                            + (f" [{nota}]" if nota else "")
                            + f" | {c}/{n} prontos"
                        )
                    except Exception as e:
                        i_f = futures[future]
                        log.append(f"   Erro no chunk {i_f+1}: {e}")
                    yield "\n".join(log), "", ""

            t_extr = int(time.time() - t_inicio)
            log.append(f"   Extracao: {t_extr//60}min{t_extr%60:02d}s")
            yield "\n".join(log), "", ""

            # Configurar cache (apenas para versao normal — resumida nao usa)
            if versao_resumida:
                cache = None
                log.append("\nModo resumo — cache nao necessario.")
            else:
                log.append("\nConfigurando cache para consolidacao...")
                yield "\n".join(log), "", ""
                cache = _proc_obter_cache(client1, model_cons)
                log[-1] = "Cache configurado." if cache else "Cache nao disponivel — usando prompt completo."
            yield "\n".join(log), "", ""

            # Consolidar
            log.append(f"\nConsolidando relatorio ({model_cons})...")
            yield "\n".join(log), "", ""
            lista = [parciais.get(i, "") for i in range(n)]
            relatorio = _proc_consolidar(client1, lista, instrucoes, cache, model_cons, versao_resumida)

        # Processos relacionados
        if pdf_relacionados:
            pdf_paths_rel = [f.name if hasattr(f, "name") else str(f) for f in pdf_relacionados]
            log.append(f"\nProcessando {len(pdf_paths_rel)} processo(s) relacionado(s)...")
            yield "\n".join(log), "", ""
            secao_rel = _proc_processar_relacionados(pdf_paths_rel, client1, client2, instrucoes, cache, model_cons)
            if secao_rel:
                relatorio = relatorio + "\n\n" + secao_rel
            log[-1] = "   Processos relacionados analisados." if secao_rel else "   Nenhuma informacao relevante nos processos relacionados."
            yield "\n".join(log), "", ""

        t_total = int(time.time() - t_inicio)
        log.append(f"\nAnalise concluida em {t_total//60}min{t_total%60:02d}s | {len(relatorio):,} chars")
        relatorio_state = relatorio
        yield "\n".join(log), relatorio, relatorio_state

    except Exception as exc:
        log.append(f"\nErro: {exc}")
        yield "\n".join(log), f"Erro:\n{exc}", ""

    finally:
        for cp, _, _, orig in todos_chunks:
            if cp != orig:
                try: os.remove(cp)
                except: pass
        for cp in _temp_comprimidos:
            try: os.remove(cp)
            except: pass


def proc_gerar_word(relatorio: str):
    if not relatorio.strip():
        return gr.update(value=None, visible=False)
    return gr.update(value=_gerar_docx(relatorio, "Analise de Processo Judicial"), visible=True)


def proc_responder(pergunta: str, relatorio: str):
    try:
        k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k1)
        return _responder_pergunta_generica(pergunta, relatorio, client, MODEL_PROC_CONS)
    except Exception as e:
        return f"Erro: {e}"



