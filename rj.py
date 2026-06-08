# -*- coding: utf-8 -*-
import os
import math
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

from report_template_rj import REPORT_TEMPLATE_RJ, SYSTEM_PROMPT_RJ
from utils import _retry, _gerar_docx, _responder_pergunta_generica, _get_clients_rj, _barra_progresso

CHUNK_MAX_PAGES_RJ    = 200
MODEL_EXTRACAO_RJ     = os.getenv("GEMINI_MODEL_EXTRACAO", "gemini-2.5-flash")
MODEL_RAPIDO_RJ       = os.getenv("GEMINI_MODEL_RAPIDO", "gemini-2.5-flash")
MODEL_PRO_RJ          = os.getenv("GEMINI_MODEL_CONSOLIDACAO", "gemini-2.5-pro")
MODEL_CONSOLIDACAO_RJ = MODEL_PRO_RJ
AVG_MIN_POR_CHUNK_RJ  = 4
MIN_CONSOLIDACAO_RJ   = 8
# MÓDULO — ANÁLISE DE RECUPERAÇÃO JUDICIAL (Teste B)
# ══════════════════════════════════════════════════════════════════════════════

_rj_cache_nome: dict[str, str] = {}
_rj_cache_lock = threading.Lock()


def _rj_dividir_pdf(path: str) -> list:
    try:
        doc = fitz.open(path)
        total = len(doc)
        if total <= CHUNK_MAX_PAGES_RJ:
            doc.close()
            return [(path, 0, total)]
        chunks = []
        for start in range(0, total, CHUNK_MAX_PAGES_RJ):
            end = min(start + CHUNK_MAX_PAGES_RJ, total)
            sub = fitz.open()
            sub.insert_pdf(doc, from_page=start, to_page=end - 1)
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            sub.save(tmp.name, garbage=4, deflate=True)
            sub.close()
            tmp.close()
            chunks.append((tmp.name, start, total))
        doc.close()
        return chunks
    except Exception:
        return [(path, 0, 0)]


def _rj_extrair_chunk(args) -> tuple:
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    pg_fim = min(offset + CHUNK_MAX_PAGES_RJ, total_pg) if total_pg else "?"

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
    PROMPT_RJ = (
        "Voce esta analisando um trecho de um processo de Recuperacao Judicial brasileiro.\n"
        "Extraia COM MAXIMO DETALHE todas as informacoes presentes nestas paginas.\n"
        "Para paginas escaneadas: aplique OCR visual completo.\n\n"
        "Cubra TODOS os itens encontrados:\n"
        "- Recuperandos (nome, CPF/CNPJ, papel)\n"
        "- Administrador Judicial\n"
        "- Advogados e escritorios (nome, OAB)\n"
        "- Datas relevantes: pedido, deferimento, AGC, PRJ, RMA\n"
        "- Status do PRJ, AGC, RMA, QGC\n"
        "- Stay period (prazo, prorrogacoes)\n"
        "- Endividamento fiscal e tributario\n"
        "- Imoveis, contratos, CCBs (numero, valor, garantia, avalistas, indices)\n"
        "- Creditos extraconcursais\n"
        "- Execucoes relacionadas (numero, partes, valor, status, andamentos)\n"
        "- Divergencias, impugnacoes e habilitacoes de credito\n"
        "- Andamentos processuais com datas (TODOS)\n"
        "Nao omita nada. Nao invente. Se nao encontrar, diga explicitamente.\n"
    )

    cabecalho = (
        f"[PARTE {idx+1}/{n_total} — paginas {pg_ini}-{pg_fim}]\n"
        "Nao omita nenhum dado mesmo que pareca repetitivo.\n"
    )
    prompt_txt = cabecalho + PROMPT_RJ
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
        resp = client.models.generate_content(model=MODEL_EXTRACAO_RJ, contents=contents)
        return resp.text

    resultado = _retry(_call)
    nota = f"{n_scan} pag. escaneadas via imagem" if n_scan else ""
    return idx, resultado, nota


def _rj_obter_cache(client, model_cons: str) -> Optional[str]:
    global _rj_cache_nome
    with _rj_cache_lock:
        if model_cons in _rj_cache_nome:
            # Valida se o cache ainda existe na API (TTL de 1h pode ter expirado)
            try:
                client.caches.get(name=_rj_cache_nome[model_cons])
                return _rj_cache_nome[model_cons]
            except Exception:
                _rj_cache_nome.pop(model_cons, None)
        try:
            cache = client.caches.create(
                model=model_cons,
                config=types.CreateCachedContentConfig(
                    display_name="invista_rj_template_v1",
                    system_instruction=SYSTEM_PROMPT_RJ,
                    contents=[types.Content(
                        role="user",
                        parts=[types.Part(text=f"TEMPLATE DO RELATORIO:\n\n{REPORT_TEMPLATE_RJ}")]
                    )],
                    ttl="3600s",
                )
            )
            _rj_cache_nome[model_cons] = cache.name
            return cache.name
        except Exception:
            return None


def _rj_merge_textos(extracoes: list) -> str:
    partes = []
    for i, txt in enumerate(extracoes):
        if txt:
            partes.append(f"{'='*60}\nPARTE {i+1}/{len(extracoes)}\n{'='*60}\n{txt}")
    return "\n\n".join(partes)


TEMPLATE_RESUMIDO_RJ = """\
MODO RESUMO PROCESSUAL — siga rigorosamente este formato curto:

A. Recuperação Judicial nº [número completo] - [Vara] - [Tribunal]

• Partes:
   ∘ Recuperando(s): [Nome] — [CPF/CNPJ]
   ∘ Administrador Judicial: [Nome] — [OAB]

• Data de distribuição: DD/MM/AAAA

• Data de deferimento do processamento: DD/MM/AAAA

• Principais andamentos: (lista cronológica completa do mais antigo ao mais recente, com data, descrição objetiva e referência Mov./fls.)
   ▪ DD/MM/AAAA — [Descrição objetiva] (Mov. X | fls. XX/XX)
   ▪ DD/MM/AAAA — [Próximo ato] (Mov. X | fls. XX/XX)
   ...

REGRAS RÍGIDAS:
- NÃO gere análise de créditos, PRJ, AGC, RMA, QGC, stay period detalhado, endividamento fiscal, garantias, contratos, índices ou qualquer outra seção.
- NÃO inclua nenhum conteúdo além das 5 seções acima.
- Use os marcadores • ∘ ▪ conforme o exemplo.
- Mantenha objetividade nos andamentos.
"""


def _rj_consolidar_secao_a(client, texto_merged: str, instrucoes: str, cache_name, model_cons: str, versao_resumida: bool = False) -> str:
    if versao_resumida:
        instrucoes_extras = f"\n\nINSTRUCOES ADICIONAIS: {instrucoes.strip()}" if instrucoes.strip() else ""
        n_partes = texto_merged.count("PARTE ") if texto_merged else 0
        prompt_full = (
            SYSTEM_PROMPT_RJ + "\n\n"
            + "Você está em MODO RESUMO. Ignore qualquer template detalhado e siga ESTRITAMENTE o formato abaixo.\n\n"
            + TEMPLATE_RESUMIDO_RJ
            + instrucoes_extras
            + f"\n\nA seguir estão {n_partes} PARTE(S) de informações extraídas do processo.\n"
            + "INSTRUÇÃO OBRIGATÓRIA DE LEITURA:\n"
            + "1. Leia CADA PARTE individualmente, do começo ao fim, sem pular nenhuma.\n"
            + "2. Extraia todos os andamentos de CADA UMA das partes (PARTE 1, PARTE 2, ..., PARTE N).\n"
            + "3. Ao final, mescle todos os andamentos em UMA única lista cronológica global.\n"
            + "4. Omitir andamentos de qualquer PARTE é ERRO GRAVE — antes de finalizar, confira se a quantidade de andamentos extraídos é compatível com a totalidade das partes.\n\n"
            + texto_merged
        )
        def _fn():
            return client.models.generate_content(
                model=model_cons, contents=[prompt_full]
            ).text
        return _retry(_fn)

    prompt = (
        "Com base nas informacoes extraidas abaixo (de multiplas partes do processo), "
        "gere APENAS a Secao A do relatorio de Recuperacao Judicial, "
        "seguindo RIGOROSAMENTE o template fornecido.\n\n"
        "Consolide informacoes duplicadas — priorize a mais completa e recente.\n\n"
        f"INFORMACOES EXTRAIDAS:\n{texto_merged}\n\n"
    )
    if instrucoes.strip():
        prompt += f"INSTRUCOES ADICIONAIS: {instrucoes.strip()}\n\n"
    prompt += (
        "Gere APENAS a Secao A — de '1. VISAO JURIDICA' ate antes de qualquer "
        "secao B.1 ou B.2. Use todos os marcadores e sub-niveis do template."
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
                with _rj_cache_lock:
                    _rj_cache_nome.pop(model_cons, None)
                conteudo = SYSTEM_PROMPT_RJ + "\n\n" + REPORT_TEMPLATE_RJ + "\n\n" + prompt
                def _fn_fallback():
                    return client.models.generate_content(
                        model=model_cons, contents=[conteudo]
                    ).text
                return _retry(_fn_fallback)
            raise
    else:
        conteudo = SYSTEM_PROMPT_RJ + "\n\n" + REPORT_TEMPLATE_RJ + "\n\n" + prompt
        def _fn():
            return client.models.generate_content(
                model=model_cons, contents=[conteudo]
            ).text
        return _retry(_fn)


def _rj_processar_relacionados(pdf_paths: list, client1, client2, instrucoes: str, cache_name, model_cons: str) -> str:
    try:
        texto_relacionados = ""
        for idx, path in enumerate(pdf_paths):
            chunks = _rj_dividir_pdf(path)
            cli = client1 if idx % 2 == 0 else client2
            for chunk_path, offset, total in chunks:
                try:
                    _, res, _ = _rj_extrair_chunk((idx, chunk_path, offset, total, len(chunks), cli))
                    if res:
                        texto_relacionados += f"\n--- {Path(path).name} ---\n{res}"
                except Exception:
                    pass

        if not texto_relacionados.strip():
            return ""

        prompt = (
            "Com base nos processos relacionados abaixo, "
            "gere secoes B.1, B.2, etc., conforme o template.\n\n"
            "REGRAS:\n"
            "- Recursos e Embargos a Execucao: integre na Secao A do processo principal\n"
            "- IDPJs, Paulianas, Embargos de Terceiro: crie secoes B proprias\n"
            "Use: 'B.1. Incidente de Desconsideracao da Personalidade n. XXX'\n\n"
            f"PROCESSOS RELACIONADOS:\n{texto_relacionados}\n\n"
        )
        if instrucoes.strip():
            prompt += f"INSTRUCOES: {instrucoes.strip()}\n\n"
        prompt += "Gere APENAS as secoes de processos relacionados (B.1, B.2, etc.)."

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
                    with _rj_cache_lock:
                        _rj_cache_nome.pop(model_cons, None)
                    conteudo = SYSTEM_PROMPT_RJ + "\n\n" + REPORT_TEMPLATE_RJ + "\n\n" + prompt
                    def _fn_fallback():
                        return client1.models.generate_content(
                            model=model_cons, contents=[conteudo]
                        ).text
                    return _retry(_fn_fallback)
                raise
        else:
            conteudo = SYSTEM_PROMPT_RJ + "\n\n" + REPORT_TEMPLATE_RJ + "\n\n" + prompt
            def _fn():
                return client1.models.generate_content(
                    model=model_cons, contents=[conteudo]
                ).text
            return _retry(_fn)
    except Exception as e:
        return f"Erro ao processar relacionados: {e}"



def rj_analisar(pdf_files, pdf_relacionados, instrucoes: str, usar_gemini_pro: bool = False, versao_resumida: bool = False):
    if not pdf_files:
        yield "Nenhum arquivo enviado.", "", ""
        return

    try:
        client1, client2 = _get_clients_rj()
    except Exception as e:
        yield f"Erro de configuracao: {e}", "", ""
        return

    pdf_paths = sorted(
        [f.name if hasattr(f, "name") else str(f) for f in pdf_files],
        key=lambda p: Path(p).name.lower()
    )
    log = []
    model_cons = MODEL_PRO_RJ if usar_gemini_pro else MODEL_RAPIDO_RJ

    log.append(f"Arquivo(s) recebido(s): {len(pdf_paths)}")
    for p in pdf_paths:
        mb = Path(p).stat().st_size / 1_048_576
        log.append(f"   · {Path(p).name} ({mb:.1f} MB)")
    log.append(f"Modelo de consolidacao: {model_cons}")
    yield "\n".join(log), "", ""

    log.append("\nInspecionando e dividindo PDFs...")
    yield "\n".join(log), "", ""

    todos_chunks = []
    for path in pdf_paths:
        nome = Path(path).name
        chunks = _rj_dividir_pdf(path)
        total_pg = chunks[0][2] if chunks else 0
        doc_tmp = fitz.open(path)
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
        log.append(f"   · {nome}: {total_pg} pag. — {info_scan}")
        if len(chunks) > 1:
            log.append(f"     Dividido em {len(chunks)} partes ({CHUNK_MAX_PAGES_RJ} pag. cada)")
        for chunk_path, offset, total in chunks:
            todos_chunks.append((chunk_path, offset, total, path))

    n = len(todos_chunks)
    n_rodadas = math.ceil(n / 2)
    tempo_est = n_rodadas * AVG_MIN_POR_CHUNK_RJ + MIN_CONSOLIDACAO_RJ
    log.append(f"\nEstimativa: ~{tempo_est} min | {n} chunk(s) · 2 chaves paralelas · {model_cons}")
    yield "\n".join(log), "", ""

    t_inicio = time.time()

    try:
        # Extracao paralela inline
        log.append(f"\nExtraindo {n} chunk(s) em paralelo (Flash 2.5 · inline)...")
        log.append(_barra_progresso(0, n))
        yield "\n".join(log), "", ""

        parciais: dict = {}

        def _worker_rj(idx):
            cp, offset, total_pg, _ = todos_chunks[idx]
            cli = client1 if idx % 2 == 0 else client2
            return _rj_extrair_chunk((idx, cp, offset, total_pg, n, cli))

        t_extr = time.time()
        with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
            futures = {ex.submit(_worker_rj, i): i for i in range(n)}
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, res, nota = future.result()
                    parciais[idx] = res
                    c = len(parciais)
                    log[-1] = _barra_progresso(c, n)
                    est_rest = (n - c) * AVG_MIN_POR_CHUNK_RJ
                    log.append(
                        f"   Chunk {idx+1}/{n} extraido"
                        + (f" [{nota}]" if nota else "")
                        + f" | {c}/{n} prontos | ~{est_rest}min restantes"
                    )
                except Exception as e:
                    i_f = futures[future]
                    log.append(f"   Erro no chunk {i_f+1}: {e}")
                yield "\n".join(log), "", ""

        t_extr_s = int(time.time() - t_extr)
        log.append(f"   Extracao total: {t_extr_s//60}min{t_extr_s%60:02d}s")
        yield "\n".join(log), "", ""

        # Merge
        log.append("\nConsolidando textos extraidos...")
        yield "\n".join(log), "", ""
        lista = [parciais.get(i, "") for i in range(n)]
        texto_merged = _rj_merge_textos(lista)
        log.append(f"   {len(texto_merged):,} caracteres de informacao extraida")
        yield "\n".join(log), "", ""

        # Cache
        log.append(f"\nConfigurando context cache ({model_cons})...")
        yield "\n".join(log), "", ""
        cache = _rj_obter_cache(client1, model_cons)
        log[-1] = "Cache configurado." if cache else "Cache nao disponivel — usando prompt completo."
        yield "\n".join(log), "", ""

        # Secao A
        secoes = []
        log.append(f"\nGerando Secao A — Recuperacao Judicial ({model_cons})...")
        log.append(f"   Aguardando {model_cons}... (pode levar alguns minutos)")
        yield "\n".join(log), "", ""

        secao_a = _rj_consolidar_secao_a(client1, texto_merged, instrucoes, cache, model_cons, versao_resumida)
        log[-1] = "   Secao A gerada."
        secoes.append(secao_a)
        yield "\n".join(log), "", ""

        # Processos relacionados
        if pdf_relacionados:
            pdf_paths_rel = [f.name if hasattr(f, "name") else str(f) for f in pdf_relacionados]
            log.append(f"\nProcessando {len(pdf_paths_rel)} processo(s) relacionado(s)...")
            yield "\n".join(log), "", ""
            secao_rel = _rj_processar_relacionados(pdf_paths_rel, client1, client2, instrucoes, cache, model_cons)
            if secao_rel and "nenhum" not in secao_rel.lower():
                log[-1] = "   Processos relacionados analisados."
                secoes.append(secao_rel)
            else:
                log[-1] = "   Nenhuma informacao relevante nos processos relacionados."
            yield "\n".join(log), "", ""

        relatorio = "\n\n".join(s for s in secoes if s)
        t_total = int(time.time() - t_inicio)
        log.append(f"\nAnalise concluida em {t_total//60}min{t_total%60:02d}s | {len(relatorio):,} chars")
        yield "\n".join(log), relatorio, relatorio

    except Exception as exc:
        log.append(f"\nErro: {exc}")
        yield "\n".join(log), f"Erro:\n{exc}", ""

    finally:
        for cp, _, _, orig in todos_chunks:
            if cp != orig:
                try: os.remove(cp)
                except: pass


def rj_gerar_word(relatorio: str):
    if not relatorio.strip():
        return gr.update(value=None, visible=False)
    return gr.update(value=_gerar_docx(relatorio, "Analise de Recuperacao Judicial"), visible=True)


def rj_responder(pergunta: str, relatorio: str):
    try:
        k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k1)
        return _responder_pergunta_generica(pergunta, relatorio, client, MODEL_CONSOLIDACAO_RJ)
    except Exception as e:
        return f"Erro: {e}"



