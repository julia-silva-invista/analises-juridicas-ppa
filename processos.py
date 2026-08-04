# -*- coding: utf-8 -*-
import os
import re
import time
import tempfile
import threading
import traceback
import concurrent.futures
from contextlib import nullcontext
from pathlib import Path
from typing import Optional

import fitz
import gradio as gr
from google.genai import types

from report_template_processos import REPORT_TEMPLATE_INSTRUCTIONS, SYSTEM_PROMPT as SYSTEM_PROMPT_PROC
from utils import (
    _retry, _gerar_docx, _responder_pergunta_generica, _get_gemini_clients,
    _executar_com_failover_gemini,
    _barra_progresso, _filtrar_arquivos_existentes,
    GEMINI_MODEL_EXTRACAO, GEMINI_MODEL_RELATORIO, GEMINI_MODEL_ESTRUTURADO, GEMINI_MODEL_QA,
)
from analysis_runtime import (
    ANALYSIS_MANAGER, cleanup_chunk, file_sha256, iter_pdf_chunks, load_chunk_cache,
    pdf_preparation_slot, queue_message, record_status, save_chunk_cache,
)
from dossie_ppa import gerar_dossie_word

CHUNK_MAX_PAGES_PROC    = 400
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
    "REFERENCIA OBRIGATORIA EM TODA INFORMACAO EXTRAIDA: cada pagina do processo tem um identificador "
    "proprio do sistema do tribunal, estampado no cabecalho/rodape/lateral da pagina — 'fls. NN' (fisico/"
    "eSAJ), 'Mov. NN' (PJe), 'Evento NN' (Eproc) ou 'ID NNNNNN' (Projudi/outro). Ao extrair QUALQUER dado "
    "abaixo (nao so andamentos), registre junto o identificador exato da pagina/movimento onde aquele dado "
    "especifico aparece — nao apenas uma vez no topo do trecho, mas ao lado de cada informacao. Se essa "
    "referencia nao for capturada agora, ela se perde e nao pode ser reconstruida depois. Nunca invente ou "
    "estime uma referencia — se nao conseguir ler o identificador da pagina (ex.: digitalizacao ilegivel), "
    "registre isso explicitamente em vez de adivinhar.\n\n"
    "Extraia e registre COM MAXIMO DETALHE todas as informacoes presentes nestas paginas:\n"
    "- Partes (nomes, CPF/CNPJ, qualidade: exequente, executado, avalista, etc.)\n"
    "- Advogados e OABs\n"
    "- Datas de todos os atos\n"
    "- Valores monetarios (exatamente como constam)\n"
    "- Indices contratuais do LASTRO/titulo executivo (contrato, CCB, CPR, CPRF, duplicata): juros "
    "remuneratorios, moratorios, correcao monetaria, multa — o INDICE OU TAXA que o contrato disciplina,\n"
    "  nao um valor calculado a partir dele.\n"
    "- Indices/taxas aplicados em cada planilha de calculo/memoria de debito encontrada (idem: o "
    "indice/taxa usado no calculo, alem do valor total resultante)\n"
    "- Identificacao dos titulos: numero da CCB, contrato, duplicata, etc.\n"
    "- Garantias e seus detalhes\n"
    "- Assinaturas e representantes\n"
    "- Citacoes, intimacoes e suas modalidades\n"
    "- Excecoes, embargos e recursos: teses e status\n"
    "- Decisoes e despachos\n"
    "- Andamentos processuais: registre todos os atos relevantes com data e referencia (Mov./fls./Evento/ID).\n"
    "  Incluir: decisoes e despachos (mesmo interlocutorios), citacoes positivas, peticoes relevantes\n"
    "  das partes, acordos e homologacoes/inadimplementos, constricoes deferidas e indeferidas,\n"
    "  planilhas de debito juntadas, penhoras no rosto dos autos, recursos, embargos, leiloes,\n"
    "  transito em julgado. Omitir apenas certidoes de expediente, citacoes negativas e juntadas\n"
    "  sem conteudo decisorio. NAO omita periodos — a consolidacao precisa de todos os fragmentos.\n\n"
    "- Penhoras no rosto dos autos: se houver requerimento de OUTROS credores (distintos do exequente)\n"
    "  pedindo reserva de valores/bens nesta execucao, registre separadamente com nome, CPF/CNPJ,\n"
    "  valor, origem do credito e status (deferido/indeferido/pendente).\n\n"
    "NAO formate como relatorio final ainda. Apenas extraia tudo com fidelidade."
)


def _proc_dividir_pdf(pdf_path: str) -> list:
    chunks = []
    try:
        with pdf_preparation_slot():
            for chunk in iter_pdf_chunks(pdf_path, max_pages=CHUNK_MAX_PAGES_PROC):
                chunks.append(chunk)
        return chunks
    except Exception:
        for chunk in chunks:
            cleanup_chunk(chunk)
        raise


def _proc_extrair_chunk(args) -> tuple:
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    try:
        with fitz.open(chunk_path) as _chunk_doc:
            pg_fim = offset + len(_chunk_doc)
    except Exception:
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
        resp = client.models.generate_content(model=GEMINI_MODEL_EXTRACAO, contents=contents)
        return resp.text

    resultado = _retry(_call)
    nota = f"{n_scan} pag. escaneadas via imagem" if n_scan else ""
    return idx, resultado, nota


def _proc_extrair_chunk_fileapi(args) -> tuple:
    """Faz upload do chunk preservado para o File API e extrai o PDF nativamente."""
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    try:
        with fitz.open(chunk_path) as _chunk_doc:
            pg_fim = offset + len(_chunk_doc)
    except Exception:
        pg_fim = min(offset + CHUNK_MAX_PAGES_PROC, total_pg) if total_pg else "?"
    source_for_upload = chunk_path
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

    try:
        arq = client.files.upload(file=upload_path)
    finally:
        if ascii_tmp:
            try: os.remove(ascii_tmp)
            except OSError: pass
    total_wait, wait_time = 0, 1
    while total_wait < 120:
        _st = getattr(arq, "state", None)
        state_name = getattr(_st, "name", None) or str(_st or "")
        if state_name in ("ACTIVE", "FAILED"):
            break
        time.sleep(wait_time)
        total_wait += wait_time
        arq = client.files.get(name=arq.name)
        wait_time = min(wait_time + 1, 4)

    _st = getattr(arq, "state", None)
    state_name = getattr(_st, "name", None) or str(_st or "")
    try:
        if state_name == "FAILED":
            raise RuntimeError(f"File API: upload do chunk {idx+1} falhou (FAILED)")

        time.sleep(0.5)
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

        def _call():
            return client.models.generate_content(model=GEMINI_MODEL_EXTRACAO, contents=contents).text

        resultado = _retry(_call, tentativas=2, espera_base=5)
    finally:
        try: client.files.delete(name=arq.name)
        except Exception: pass

    return idx, resultado, "File API"


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


def _proc_consolidar(client, parciais: list, instrucoes: str, cache_name, model_cons: str, versao_resumida: bool = False, nomes_arquivos: list = None) -> str:
    multi_arquivo = bool(nomes_arquivos) and len(set(nomes_arquivos)) > 1
    blocos = "\n\n".join(
        f"{'='*60}\nPARTE {i+1}/{len(parciais)}"
        + (f" — arquivo: {nomes_arquivos[i]}" if multi_arquivo else "")
        + f"\n{'='*60}\n{p}"
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


def _proc_consolidar_com_failover(
    clients: list,
    parciais: list,
    instrucoes: str,
    cache_name,
    model_cons: str,
    versao_resumida: bool = False,
    nomes_arquivos: list = None,
    log: list = None,
) -> str:
    def _ao_falhar(indice, proximo, _exc):
        if log is not None:
            log.append(
                f"   Credencial Gemini {indice + 1} recusada ou indisponível; "
                f"tentando a credencial {proximo + 1}."
            )

    return _executar_com_failover_gemini(
        clients,
        lambda client, indice: _proc_consolidar(
            client,
            parciais,
            instrucoes,
            cache_name if indice == 0 else None,
            model_cons,
            versao_resumida,
            nomes_arquivos,
        ),
        ao_falhar=_ao_falhar,
    )


def _proc_processar_relacionados(pdf_paths: list, clients: list, instrucoes: str, cache_name,
                                  model_cons: str, runtime_job=None) -> str:
    """Extrai e consolida processos relacionados, gerando seções B.1, B.2, etc."""
    try:
        todos_chunks = []
        hashes_arquivos = {}
        for path in pdf_paths:
            hashes_arquivos[path] = file_sha256(path)
            for chunk in _proc_dividir_pdf(path):
                todos_chunks.append((
                    chunk.path, chunk.start, chunk.total_pages, path,
                    chunk.preparation_note,
                ))

        resultados = {}

        def _worker_rel(i):
            chunk_path, offset, total, original, preparation_note = todos_chunks[i]
            try:
                with fitz.open(chunk_path) as _chunk_doc:
                    chunk_end = offset + len(_chunk_doc)
            except Exception:
                chunk_end = min(offset + CHUNK_MAX_PAGES_PROC, total)
            source_hash = hashes_arquivos[original]
            cached = load_chunk_cache(
                source_hash, offset, chunk_end, GEMINI_MODEL_EXTRACAO, "processo_relacionados"
            )
            if cached is not None:
                return i, cached, "cache por arquivo/páginas"
            slot = runtime_job.worker_slot() if runtime_job else nullcontext()
            with slot:
                record_status(
                    runtime_job, "extraindo_chunk", arquivo=Path(original).name,
                    paginas=f"{offset + 1}-{chunk_end}", chunk=i + 1,
                )
                trocas = []

                def _extrair_com(client, _indice):
                    try:
                        return _retry(
                            lambda: _proc_extrair_chunk_fileapi(
                                (i, chunk_path, offset, total, len(todos_chunks), client)
                            ),
                            tentativas=2, espera_base=15,
                        )
                    except Exception as exc:
                        msg = str(exc)
                        if any(code in msg for code in ["400", "INVALID_ARGUMENT", "403", "PERMISSION_DENIED"]):
                            ri, text, note = _proc_extrair_chunk(
                                (i, chunk_path, offset, total, len(todos_chunks), client)
                            )
                            return ri, text, (note + " | " if note else "") + "fallback inline"
                        raise

                result = _executar_com_failover_gemini(
                    clients,
                    _extrair_com,
                    indice_inicial=i % len(clients),
                    ao_falhar=lambda atual, proximo, _exc: trocas.append(
                        f"credencial Gemini {atual + 1}→{proximo + 1}"
                    ),
                )
                result = (
                    result[0], result[1],
                    " | ".join(filter(None, [result[2], preparation_note, *trocas])),
                )
            save_chunk_cache(
                source_hash, offset, chunk_end, GEMINI_MODEL_EXTRACAO,
                "processo_relacionados", result[1],
            )
            return result

        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as executor:
            futures = {executor.submit(_worker_rel, i): i for i in range(len(todos_chunks))}
            for future in concurrent.futures.as_completed(futures):
                i = futures[future]
                try:
                    _idx, text, _note = future.result()
                    resultados[i] = text
                except Exception:
                    resultados[i] = ""
                finally:
                    chunk_path, _offset, _total, original, _preparation_note = todos_chunks[i]
                    if chunk_path != original:
                        try: os.remove(chunk_path)
                        except OSError: pass

        texto_relacionados = ""
        for i, (_chunk_path, _offset, _total, original, _preparation_note) in enumerate(todos_chunks):
            if resultados.get(i):
                texto_relacionados += f"\n--- {Path(original).name} ---\n{resultados[i]}"

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

        def _consolidar_com(client, indice):
            if cache_name and indice == 0:
                contents = [types.Content(role="user", parts=[types.Part(text=prompt)])]
                config = types.GenerateContentConfig(cached_content=cache_name)
                try:
                    return _retry(lambda: client.models.generate_content(
                        model=model_cons, contents=contents, config=config
                    ).text)
                except Exception as e:
                    msg = str(e)
                    if "PERMISSION_DENIED" in msg or "CachedContent not found" in msg or "403" in msg:
                        with _proc_cache_lock:
                            _proc_cache_nome.pop(model_cons, None)
                    else:
                        raise
            prompt_full = SYSTEM_PROMPT_PROC + "\n\n" + REPORT_TEMPLATE_INSTRUCTIONS + "\n\n" + prompt
            return _retry(lambda: client.models.generate_content(
                model=model_cons, contents=[prompt_full]
            ).text)

        return _executar_com_failover_gemini(clients, _consolidar_com)
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


def _erro_completo_relatorio_proc(exc: Exception) -> str:
    """Texto de erro pra aba Relatorio: traceback completo, nao so str(exc) — o Gradio
    (sem show_error=True) nao mostra nenhum detalhe do erro, entao a unica forma da
    usuaria ver o motivo real de um "Erro" vermelho e a propria funcao escrever isso
    num output visivel. Mesmo padrao ja usado em rj.py."""
    return "❌ ERRO — a análise foi interrompida.\n\n" + traceback.format_exc()


def proc_analisar(pdf_files, pdf_relacionados, instrucoes: str, versao_resumida: bool = False):
    runtime_job = ANALYSIS_MANAGER.create("processo")
    try:
        while True:
            active, position = runtime_job.try_activate()
            if active:
                break
            yield queue_message(position), "", "", ""
            runtime_job.wait_for_change(2.0)
        record_status(runtime_job, "iniciando")
        yield from _proc_analisar_impl(
            pdf_files, pdf_relacionados, instrucoes, versao_resumida,
            runtime_job=runtime_job,
        )
    except Exception as exc:
        yield "Erro inesperado — veja o traceback completo na aba Relatório.", _erro_completo_relatorio_proc(exc), "", ""
    finally:
        record_status(runtime_job, "finalizada")
        runtime_job.close()


def _proc_analisar_impl(pdf_files, pdf_relacionados, instrucoes: str, versao_resumida: bool = False,
                        runtime_job=None):
    if not pdf_files:
        yield "Nenhum arquivo enviado.", "", "", ""
        return

    yield "Iniciando analise de processo judicial...", "", "", ""

    try:
        clients = _get_gemini_clients()
    except Exception as e:
        yield f"Erro de configuracao: {e}", "", "", ""
        return

    def _nat_key(p):
        return [int(x) if x.isdigit() else x.lower()
                for x in re.split(r"(\d+)", Path(p).name)]

    pdf_paths = sorted(
        [f.name if hasattr(f, "name") else str(f) for f in pdf_files],
        key=_nat_key
    )
    log = []
    pdf_paths = _filtrar_arquivos_existentes(pdf_paths, log)
    if not pdf_paths:
        log.append("\nNenhum dos arquivos enviados pôde ser lido — tente reenviar.")
        yield "\n".join(log), "", "", ""
        return

    relatorio_state = ""
    model_cons = GEMINI_MODEL_RELATORIO

    log.append(f"Arquivo(s) recebido(s): {len(pdf_paths)}")
    for p in pdf_paths:
        mb = Path(p).stat().st_size / 1_048_576
        log.append(f"   · {Path(p).name} ({mb:.1f} MB)")
    if len(pdf_paths) > 1:
        log.append(f"   → {len(pdf_paths)} arquivos — o modelo verificara se sao fragmentos do mesmo processo ou processos distintos")
    if instrucoes.strip():
        log.append("Instrucoes adicionais recebidas.")
    log.append(f"Modelo de consolidacao: {model_cons}")
    yield "\n".join(log), "", "", ""

    # Inspecionar e dividir
    log.append("\nInspecionando e dividindo PDFs sem recompressão de imagens...")
    record_status(runtime_job, "dividindo_pdfs")
    yield "\n".join(log), "", "", ""

    todos_chunks = []
    hashes_arquivos = {}
    for path in pdf_paths:
        nome = Path(path).name
        mb_orig = Path(path).stat().st_size / 1_048_576
        hashes_arquivos[path] = file_sha256(path)
        record_status(runtime_job, "dividindo_pdf", arquivo=nome, tamanho_mb=round(mb_orig, 1))
        doc_tmp = fitz.open(path)
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
        chunks = _proc_dividir_pdf(path)
        log.append(f"   · {nome}: {total_pg} pag. — {info_scan}")
        if len(chunks) > 1:
            log.append(f"     Dividido em {len(chunks)} partes de até {CHUNK_MAX_PAGES_PROC} pag. e 45 MB")
        for chunk in chunks:
            todos_chunks.append((
                chunk.path, chunk.start, chunk.total_pages, path,
                chunk.preparation_note,
            ))
            log.append(
                f"     Páginas {chunk.page_start}-{chunk.page_end}: "
                f"{chunk.preparation_note}"
            )

    n = len(todos_chunks)
    chunks_total_mb = sum(Path(c[0]).stat().st_size for c in todos_chunks) / 1_048_576
    yield "\n".join(log), "", "", ""

    t_inicio = time.time()
    cache = None
    texto_merged = ""  # texto OCR completo — fonte para o dossiê PPA

    try:
        workers_iniciais = ANALYSIS_MANAGER.worker_cap() if runtime_job else 6
        record_status(runtime_job, "extraindo", chunks_total=n, chunks_concluidos=0)
        log.append(
            f"\nExtração em chunks ({n} chunk(s) · {chunks_total_mb:.0f} MB · até "
            f"{workers_iniciais} worker(s) nesta análise / 6 globais · File API)..."
        )
        log.append(_barra_progresso(0, n))
        yield "\n".join(log), "", "", ""

        parciais: dict = {}

        def _worker_proc(idx):
            cp, offset, total_pg, original, preparation_note = todos_chunks[idx]
            try:
                with fitz.open(cp) as _chunk_doc:
                    chunk_end = offset + len(_chunk_doc)
            except Exception:
                chunk_end = min(offset + CHUNK_MAX_PAGES_PROC, total_pg)
            source_hash = hashes_arquivos[original]
            cached = load_chunk_cache(
                source_hash, offset, chunk_end, GEMINI_MODEL_EXTRACAO, "processo_principal"
            )
            if cached is not None:
                return idx, cached, "cache por arquivo/páginas"
            slot = runtime_job.worker_slot() if runtime_job else nullcontext()
            with slot:
                record_status(
                    runtime_job, "extraindo_chunk", arquivo=Path(original).name,
                    paginas=f"{offset + 1}-{chunk_end}", chunk=idx + 1,
                )
                trocas = []

                def _extrair_com(client, _indice):
                    try:
                        return _retry(
                            lambda: _proc_extrair_chunk_fileapi((idx, cp, offset, total_pg, n, client)),
                            tentativas=2, espera_base=15,
                        )
                    except Exception as e:
                        msg_e = str(e)
                        if any(c in msg_e for c in ["400", "INVALID_ARGUMENT", "403", "PERMISSION_DENIED"]):
                            ri, res, nota = _proc_extrair_chunk((idx, cp, offset, total_pg, n, client))
                            return ri, res, (nota + " | " if nota else "") + "fallback inline"
                        raise

                result = _executar_com_failover_gemini(
                    clients,
                    _extrair_com,
                    indice_inicial=idx % len(clients),
                    ao_falhar=lambda atual, proximo, _exc: trocas.append(
                        f"credencial Gemini {atual + 1}→{proximo + 1}"
                    ),
                )
                result = (
                    result[0], result[1],
                    " | ".join(filter(None, [result[2], preparation_note, *trocas])),
                )
            save_chunk_cache(
                source_hash, offset, chunk_end, GEMINI_MODEL_EXTRACAO,
                "processo_principal", result[1],
            )
            return result

        with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
            futures = {ex.submit(_worker_proc, i): i for i in range(n)}
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, res, nota = future.result()
                    parciais[idx] = res
                    c = len(parciais)
                    record_status(
                        runtime_job, "extraindo", chunks_total=n, chunks_concluidos=c,
                        arquivo=Path(todos_chunks[idx][3]).name,
                    )
                    log[-1] = _barra_progresso(c, n)
                    log.append(
                        f"   Chunk {idx+1}/{n} extraido"
                        + (f" [{nota}]" if nota else "")
                        + f" | {c}/{n} prontos"
                    )
                except Exception as e:
                    i_f = futures[future]
                    log.append(f"   Erro no chunk {i_f+1}: {e}")
                finally:
                    # Libera o chunk do disco assim que termina (sucesso ou erro), em
                    # vez de esperar a analise inteira acabar — reduz pico de disco.
                    i_f = futures[future]
                    cp_done, _, _, orig_done, _preparation_note = todos_chunks[i_f]
                    if cp_done != orig_done:
                        try: os.remove(cp_done)
                        except Exception: pass
                yield "\n".join(log), "", "", ""

        t_extr = int(time.time() - t_inicio)
        log.append(f"   Extracao: {t_extr//60}min{t_extr%60:02d}s")
        yield "\n".join(log), "", "", ""

        # Configurar cache (apenas para versao normal — resumida nao usa)
        if versao_resumida:
            cache = None
            log.append("\nModo resumo — cache nao necessario.")
        else:
            log.append("\nConfigurando cache para consolidacao...")
            yield "\n".join(log), "", "", ""
            cache = _proc_obter_cache(clients[0], model_cons)
            log[-1] = "Cache configurado." if cache else "Cache nao disponivel — usando prompt completo."
        yield "\n".join(log), "", "", ""

        record_status(runtime_job, "consolidando", chunks_total=n, chunks_concluidos=len(parciais))
        # Consolidar
        log.append(f"\nConsolidando relatorio ({model_cons})...")
        yield "\n".join(log), "", "", ""
        lista = [parciais.get(i, "") for i in range(n)]
        nomes_por_chunk = [Path(todos_chunks[i][3]).name for i in range(n)]
        multi_arquivo = len(set(nomes_por_chunk)) > 1
        texto_merged = "\n\n".join(
            f"[PARTE {i+1}]" + (f" — arquivo: {nomes_por_chunk[i]}" if multi_arquivo else "") + f"\n{t}"
            for i, t in enumerate(lista) if t and t.strip()
        )
        relatorio = _proc_consolidar_com_failover(
            clients, lista, instrucoes, cache, model_cons,
            versao_resumida, nomes_por_chunk, log,
        )

        # Processos relacionados
        if pdf_relacionados:
            pdf_paths_rel = [f.name if hasattr(f, "name") else str(f) for f in pdf_relacionados]
            log.append(f"\nProcessando {len(pdf_paths_rel)} processo(s) relacionado(s)...")
            yield "\n".join(log), "", "", ""
            secao_rel = _proc_processar_relacionados(
                pdf_paths_rel, clients, instrucoes, cache, model_cons,
                runtime_job=runtime_job,
            )
            if secao_rel:
                relatorio = relatorio + "\n\n" + secao_rel
            log[-1] = "   Processos relacionados analisados." if secao_rel else "   Nenhuma informacao relevante nos processos relacionados."
            yield "\n".join(log), "", "", ""

        t_total = int(time.time() - t_inicio)
        log.append(f"\nAnalise concluida em {t_total//60}min{t_total%60:02d}s | {len(relatorio):,} chars")
        relatorio_state = relatorio
        yield "\n".join(log), relatorio, relatorio_state, texto_merged

    except Exception as exc:
        log.append("\nErro — veja o traceback completo na aba Relatório.")
        yield "\n".join(log), _erro_completo_relatorio_proc(exc), "", ""

    finally:
        for cp, _, _, orig, _preparation_note in todos_chunks:
            if cp != orig:
                try: os.remove(cp)
                except: pass


def proc_gerar_word(relatorio: str):
    if not relatorio.strip():
        return gr.update(value=None, visible=False)
    return gr.update(value=_gerar_docx(relatorio, "Analise de Processo Judicial"), visible=True)


def proc_gerar_dossie(relatorio: str, extracao: str = ""):
    # Prioriza o texto OCR completo; o relatório resumido omite campos do dossiê
    fonte = extracao.strip() if extracao and extracao.strip() else (relatorio or "").strip()
    if not fonte:
        yield gr.update(value=None, visible=False), "Gere uma análise primeiro."
        return
    yield gr.update(visible=False), "⏳ Gerando dossiê PPA (extraindo dados via IA)..."
    try:
        # O dossiê exige síntese jurídica estruturada e completa; usa o modelo
        # de consolidação para reduzir omissões e preservar a fundamentação.
        caminho = _executar_com_failover_gemini(
            _get_gemini_clients(),
            lambda client, _indice: gerar_dossie_word(
                extracao or "", relatorio or "", client, GEMINI_MODEL_ESTRUTURADO
            ),
        )
        yield gr.update(value=caminho, visible=True), "✅ Dossiê gerado — clique no arquivo para baixar."
    except Exception:
        yield gr.update(value=None, visible=False), "❌ Erro ao gerar dossiê:\n\n" + traceback.format_exc()


def proc_responder(pergunta: str, relatorio: str):
    try:
        return _executar_com_failover_gemini(
            _get_gemini_clients(),
            lambda client, _indice: _responder_pergunta_generica(
                pergunta, relatorio, client, GEMINI_MODEL_QA
            ),
        )
    except Exception:
        return "❌ Erro:\n\n" + traceback.format_exc()



