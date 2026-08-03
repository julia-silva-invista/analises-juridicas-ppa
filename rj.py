# -*- coding: utf-8 -*-
import os
import re
import json
import math
import time
import tempfile
import threading
import concurrent.futures
from pathlib import Path
from typing import Optional
from collections import defaultdict

import fitz
import gradio as gr
from google import genai
from google.genai import types

from report_template_rj import REPORT_TEMPLATE_RJ, SYSTEM_PROMPT_RJ
from utils import _retry, _gerar_docx, _responder_pergunta_generica, _get_clients_rj, _barra_progresso, _comprimir_pdf, _comprimir_pdf_limite
from checklist_rj import gerar_checklist_rj, gerar_checklist_creditos

CHUNK_MAX_PAGES_RJ    = 400
MODEL_EXTRACAO_RJ     = os.getenv("GEMINI_MODEL_EXTRACAO", "gemini-2.5-flash")
MODEL_RAPIDO_RJ       = os.getenv("GEMINI_MODEL_RAPIDO", "gemini-2.5-flash")
MODEL_PRO_RJ          = os.getenv("GEMINI_MODEL_CONSOLIDACAO", "gemini-2.5-pro")
MODEL_CONSOLIDACAO_RJ = MODEL_PRO_RJ
AVG_MIN_POR_CHUNK_RJ  = 3
MIN_CONSOLIDACAO_RJ   = 8
COMPRESSAO_PRE_MB_RJ  = 50   # só pré-comprime se puder caber no modo direto
# MÓDULO — ANÁLISE DE RECUPERAÇÃO JUDICIAL (Teste B)
# ══════════════════════════════════════════════════════════════════════════════

_rj_cache_nome: dict[str, str] = {}
_rj_cache_lock = threading.Lock()

PROMPT_RJ = (
    "Voce esta analisando um trecho de um processo de Recuperacao Judicial brasileiro.\n"
    "IMPORTANTE: este arquivo pode ser um fragmento de um processo maior dividido em multiplos PDFs "
    "pelo tribunal, OU pode ser um processo independente. Extraia o numero do processo se disponivel — "
    "isso sera usado para detectar continuidade na consolidacao. Trate cada fragmento como parte de "
    "um documento continuo; a consolidacao final verificara se sao o mesmo processo ou processos distintos.\n"
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


def _rj_extrair_chunk_fileapi(args) -> tuple:
    """Comprime chunk, faz upload para o File API e extrai via leitura nativa de PDF."""
    idx, chunk_path, offset, total_pg, n_total, client = args
    pg_ini = offset + 1
    pg_fim = min(offset + CHUNK_MAX_PAGES_RJ, total_pg) if total_pg else "?"

    # Comprimir chunk antes do upload (acontece em paralelo nos workers)
    comp_chunk, orig_mb, comp_mb = _comprimir_pdf_limite(chunk_path, max_mb=40.0)
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
        _st = getattr(arq, "state", None)
        state_name = getattr(_st, "name", None) or str(_st or "")
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

    _st = getattr(arq, "state", None)
    state_name = getattr(_st, "name", None) or str(_st or "")
    if state_name == "FAILED":
        raise RuntimeError(f"File API: upload do chunk {idx+1} falhou (FAILED)")

    time.sleep(0.5)

    prompt = (
        f"[PARTE {idx+1}/{n_total} — paginas {pg_ini}-{pg_fim}]\n"
        "Nao omita nenhum dado mesmo que pareca repetitivo.\n\n"
        + PROMPT_RJ
    )
    mime = getattr(arq, "mime_type", None) or "application/pdf"
    contents = [types.Content(role="user", parts=[
        types.Part(text=prompt),
        types.Part(file_data=types.FileData(file_uri=arq.uri, mime_type=mime)),
    ])]

    def _call():
        return client.models.generate_content(model=MODEL_EXTRACAO_RJ, contents=contents).text

    resultado = _retry(_call, tentativas=2, espera_base=5)

    try: client.files.delete(name=arq.name)
    except: pass

    return idx, resultado, f"File API | {comp_nota}"


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

    n_partes_cons = texto_merged.count("PARTE ") if texto_merged else 0
    _aviso_multiplos_rj = (
        f"ATENCAO: foram carregados fragmentos de {n_partes_cons} parte(s). "
        "Verifique se compartilham o mesmo numero de processo de Recuperacao Judicial — "
        "se sim, trate como UM UNICO PROCESSO CONTINUO. "
        "Se forem processos distintos, gere secoes separadas para cada um.\n\n"
    ) if n_partes_cons > 1 else ""
    prompt = (
        "Com base nas informacoes extraidas abaixo (de multiplas partes do processo), "
        "gere APENAS a Secao A do relatorio de Recuperacao Judicial, "
        "seguindo RIGOROSAMENTE o template fornecido.\n\n"
        + _aviso_multiplos_rj
        + "Consolide informacoes duplicadas — priorize a mais completa e recente.\n\n"
        + f"INFORMACOES EXTRAIDAS:\n{texto_merged}\n\n"
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


def _rj_processar_relacionados(pdf_paths: list, client1, client2, instrucoes: str, cache_name, model_cons: str,
                                standalone: bool = False):
    """Extrai e consolida os 'processos relacionados'. GERADOR — durante a extração, produz
    ("progress", linha_de_log) a cada trecho concluído (qual processo, X/N trechos); ao final,
    produz exatamente um ("done", (secao_texto, texto_bruto, avisos)).

    texto_bruto é o texto extraído página-a-página (antes da consolidação) — quem chama deve
    somá-lo à fonte usada pelos checklists/Excel de credores, para que dados de execuções que só
    aparecem nos relacionados (não repetidos no PDF principal de RJ) não se percam.

    standalone=True é usado quando NÃO há PDF principal de RJ: os processos abaixo são tratados
    como processos PRINCIPAIS (não subordinados a nenhuma RJ), reaproveitando o mesmo template de
    'Processos Relacionados' já embutido em SYSTEM_PROMPT_RJ/REPORT_TEMPLATE_RJ.
    """
    avisos: list = []
    try:
        todos_chunks_rel = []
        for path in pdf_paths:
            for chunk_path, offset, total in _rj_dividir_pdf(path):
                todos_chunks_rel.append((chunk_path, offset, total, path))

        n_rel = len(todos_chunks_rel)
        if n_rel == 0:
            yield "done", ("", "", avisos)
            return

        nomes_arquivos = sorted({Path(p).name for _, _, _, p in todos_chunks_rel})
        yield "progress", (
            f"   {n_rel} trecho(s) a extrair de {len(pdf_paths)} processo(s): "
            + ", ".join(nomes_arquivos)
        )

        resultados: dict = {}

        def _worker_rel(i):
            cp, offset, total_pg, _ = todos_chunks_rel[i]
            cli = client1 if i % 2 == 0 else client2
            try:
                return _retry(
                    lambda: _rj_extrair_chunk_fileapi((i, cp, offset, total_pg, n_rel, cli)),
                    tentativas=2, espera_base=15
                )
            except Exception as e:
                msg_e = str(e)
                if any(c in msg_e for c in ["400", "INVALID_ARGUMENT", "403", "PERMISSION_DENIED"]):
                    ri, res, nota = _rj_extrair_chunk((i, cp, offset, total_pg, n_rel, cli))
                    return ri, res, (nota + " | " if nota else "") + "fallback inline"
                raise

        concluidos = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
            futures = {ex.submit(_worker_rel, i): i for i in range(n_rel)}
            for future in concurrent.futures.as_completed(futures):
                i_f = futures[future]
                nome_arq = Path(todos_chunks_rel[i_f][3]).name
                concluidos += 1
                try:
                    idx, res, nota = future.result()
                    resultados[idx] = res
                    yield "progress", (
                        f"   {nome_arq} — trecho {i_f+1}/{n_rel} extraido"
                        + (f" [{nota}]" if nota else "")
                        + f" | {concluidos}/{n_rel} prontos"
                    )
                except Exception as e:
                    avisos.append(f"Aviso: nao foi possivel extrair um trecho de '{nome_arq}' ({e}) — pode faltar informacao desse processo.")
                    yield "progress", f"   {nome_arq} — trecho {i_f+1}/{n_rel}: erro ({e}) | {concluidos}/{n_rel} concluidos"

        for cp, _off, _tot, orig in todos_chunks_rel:
            if cp != orig:
                try: os.remove(cp)
                except: pass

        texto_relacionados = ""
        for i, (_cp, _offset, _total_pg, orig_path) in enumerate(todos_chunks_rel):
            res = resultados.get(i, "")
            if res:
                texto_relacionados += f"\n--- {Path(orig_path).name} ---\n{res}"

        if not texto_relacionados.strip():
            yield "done", ("", "", avisos)
            return

        yield "progress", f"   Consolidando {n_rel} trecho(s) extraido(s) dos processos relacionados..."

        instrucao_anti_omissao = (
            "INSTRUCAO OBRIGATORIA DE LEITURA:\n"
            "1. Leia cada processo/parte individualmente, do comeco ao fim, sem pular nenhum.\n"
            "2. Nao omita nenhuma execucao, valor, garantia, parte ou andamento relevante de qualquer processo.\n"
            f"3. Antes de finalizar, confira se voce cobriu os {len(pdf_paths)} processo(s) fornecidos abaixo.\n\n"
        )

        if standalone:
            prompt = (
                "Os processos abaixo NAO tem nenhuma Recuperacao Judicial principal associada — cada um "
                "deve ser tratado como um processo PRINCIPAL, nao subordinado a nenhum outro. Gere a Secao "
                "1 (VISAO JURIDICA) do relatorio, com uma subsecao por processo principal (A., B., C., ...) "
                "e por incidente/recurso relacionado a ele (A.1., A.2., ...), seguindo RIGOROSAMENTE o "
                "template de 'Processos Relacionados' fornecido (mesma estrutura, marcadores e nivel de "
                "detalhe usados para execucoes, embargos, IDPJs, acoes paulianas etc.).\n\n"
                + instrucao_anti_omissao
                + f"PROCESSOS ({len(pdf_paths)} arquivo(s)):\n{texto_relacionados}\n\n"
            )
        else:
            prompt = (
                "Com base nos processos relacionados abaixo, "
                "gere secoes B.1, B.2, etc., conforme o template.\n\n"
                "REGRAS:\n"
                "- Recursos e Embargos a Execucao: integre na Secao A do processo principal\n"
                "- IDPJs, Paulianas, Embargos de Terceiro: crie secoes B proprias\n"
                "Use: 'B.1. Incidente de Desconsideracao da Personalidade n. XXX'\n\n"
                + instrucao_anti_omissao
                + f"PROCESSOS RELACIONADOS ({len(pdf_paths)} arquivo(s)):\n{texto_relacionados}\n\n"
            )
        if instrucoes.strip():
            prompt += f"INSTRUCOES: {instrucoes.strip()}\n\n"
        prompt += (
            "Gere o relatorio completo (Secao 1 apenas), sem inventar informacao."
            if standalone else
            "Gere APENAS as secoes de processos relacionados (B.1, B.2, etc.)."
        )

        if cache_name:
            contents = [types.Content(role="user", parts=[types.Part(text=prompt)])]
            config = types.GenerateContentConfig(cached_content=cache_name)
            def _fn():
                return client1.models.generate_content(
                    model=model_cons, contents=contents, config=config
                ).text
            try:
                secao = _retry(_fn)
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
                    secao = _retry(_fn_fallback)
                else:
                    raise
        else:
            conteudo = SYSTEM_PROMPT_RJ + "\n\n" + REPORT_TEMPLATE_RJ + "\n\n" + prompt
            def _fn():
                return client1.models.generate_content(
                    model=model_cons, contents=[conteudo]
                ).text
            secao = _retry(_fn)

        yield "done", (secao, texto_relacionados, avisos)
    except Exception as e:
        yield "done", (f"Erro ao processar relacionados: {e}", "", avisos)


def _rj_analisar_somente_relacionados(pdf_relacionados, instrucoes: str, usar_gemini_pro: bool = False):
    """Analisa só 'processos relacionados' (execuções/ações avulsas), sem PDF principal de RJ.

    Gera um relatório objetivo por processo (reaproveitando o template de 'Processos
    Relacionados') e devolve o texto bruto consolidado como fonte para os checklists/Excel de
    credores — permite gerar o Checklist de Créditos sem exigir upload de uma RJ.
    """
    yield "Iniciando analise dos processos relacionados (sem RJ principal)...", "", "", ""

    try:
        client1, client2 = _get_clients_rj()
    except Exception as e:
        yield f"Erro de configuracao: {e}", "", "", ""
        return

    model_cons = MODEL_PRO_RJ if usar_gemini_pro else MODEL_RAPIDO_RJ
    pdf_paths_rel = [f.name if hasattr(f, "name") else str(f) for f in pdf_relacionados]

    log = [f"Processo(s) relacionado(s) recebido(s): {len(pdf_paths_rel)}"]
    for p in pdf_paths_rel:
        mb = Path(p).stat().st_size / 1_048_576
        log.append(f"   · {Path(p).name} ({mb:.1f} MB)")
    log.append(f"Modelo de consolidacao: {model_cons}")
    yield "\n".join(log), "", "", ""

    t_inicio = time.time()
    log.append("\nExtraindo processos relacionados (paralelo · File API)...")
    yield "\n".join(log), "", "", ""

    try:
        secao_rel, texto_bruto, avisos = "", "", []
        for kind, payload in _rj_processar_relacionados(
            pdf_paths_rel, client1, client2, instrucoes, cache_name=None, model_cons=model_cons,
            standalone=True,
        ):
            if kind == "progress":
                log.append(payload)
                yield "\n".join(log), "", "", ""
            else:
                secao_rel, texto_bruto, avisos = payload

        for aviso in avisos:
            log.append(f"   {aviso}")
        if avisos:
            yield "\n".join(log), "", "", ""

        if not texto_bruto.strip():
            log.append("\nNenhuma informacao relevante extraida dos processos relacionados.")
            yield "\n".join(log), "", "", ""
            return

        relatorio = secao_rel or "Nenhuma informacao relevante extraida dos processos relacionados."
        t_total = int(time.time() - t_inicio)
        log.append(f"\nAnalise concluida em {t_total//60}min{t_total%60:02d}s | {len(relatorio):,} chars")
        yield "\n".join(log), relatorio, relatorio, texto_bruto

    except Exception as exc:
        log.append(f"\nErro: {exc}")
        yield "\n".join(log), f"Erro:\n{exc}", "", ""



def rj_analisar(pdf_files, pdf_relacionados, instrucoes: str = "", usar_gemini_pro: bool = False, versao_resumida: bool = False):
    if not pdf_files and not pdf_relacionados:
        yield "Nenhum arquivo enviado.", "", "", ""
        return

    if not pdf_files:
        # Sem PDF principal de RJ: analisa só os processos relacionados (execuções/ações
        # avulsas), num fluxo mais leve — sem Seção A (que pressupõe uma RJ principal).
        yield from _rj_analisar_somente_relacionados(pdf_relacionados, instrucoes, usar_gemini_pro)
        return

    yield "Iniciando analise de Recuperacao Judicial...", "", "", ""

    try:
        client1, client2 = _get_clients_rj()
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
    model_cons = MODEL_PRO_RJ if usar_gemini_pro else MODEL_RAPIDO_RJ

    log.append(f"Arquivo(s) recebido(s): {len(pdf_paths)}")
    for p in pdf_paths:
        mb = Path(p).stat().st_size / 1_048_576
        log.append(f"   · {Path(p).name} ({mb:.1f} MB)")
    if len(pdf_paths) > 1:
        log.append(f"   → {len(pdf_paths)} arquivos — o modelo verificara se sao fragmentos do mesmo processo ou processos distintos")
    log.append(f"Modelo de consolidacao: {model_cons}")
    yield "\n".join(log), "", "", ""

    log.append("\nInspecionando e dividindo PDFs...")
    yield "\n".join(log), "", "", ""

    todos_chunks = []
    _temp_comprimidos: list = []
    for path in pdf_paths:
        nome = Path(path).name
        mb_orig = Path(path).stat().st_size / 1_048_576

        path_proc = path
        if 10 <= mb_orig <= COMPRESSAO_PRE_MB_RJ:
            # Arquivo pequeno: pré-comprime pois pode caber no modo direto
            log.append(f"   Comprimindo {nome} ({mb_orig:.0f} MB)...")
            yield "\n".join(log), "", "", ""
            comp_path, orig_mb, comp_mb = _comprimir_pdf(path)
            if comp_path != path:
                reducao = (1 - comp_mb / orig_mb) * 100
                log[-1] = f"   {nome}: {orig_mb:.0f} MB → {comp_mb:.0f} MB (-{reducao:.0f}%)"
                yield "\n".join(log), "", "", ""
                _temp_comprimidos.append(comp_path)
                path_proc = comp_path
        elif mb_orig > COMPRESSAO_PRE_MB_RJ:
            # Arquivo grande: vai direto ao chunked; compressao por chunk em paralelo
            log.append(f"   {nome} ({mb_orig:.0f} MB) — compressao por chunk (paralelo)")
            yield "\n".join(log), "", "", ""

        chunks = _rj_dividir_pdf(path_proc)
        total_pg = chunks[0][2] if chunks else 0
        doc_tmp = fitz.open(path_proc)
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
    n_rodadas = math.ceil(n / 4)
    tempo_est = n_rodadas * AVG_MIN_POR_CHUNK_RJ + MIN_CONSOLIDACAO_RJ
    log.append(f"\nEstimativa: ~{tempo_est} min | {n} chunk(s) · 4 workers · File API · {model_cons}")
    yield "\n".join(log), "", "", ""

    t_inicio = time.time()

    texto_merged = ""  # disponivel no finally para o Excel de credores

    try:
        # Extração paralela via File API — 4 workers
        log.append(f"\nExtraindo {n} chunk(s) em paralelo (4 workers · File API)...")
        log.append(_barra_progresso(0, n))
        yield "\n".join(log), "", "", ""

        parciais: dict = {}

        def _worker_rj(idx):
            cp, offset, total_pg, _ = todos_chunks[idx]
            cli = client1 if idx % 2 == 0 else client2
            try:
                return _retry(
                    lambda: _rj_extrair_chunk_fileapi((idx, cp, offset, total_pg, n, cli)),
                    tentativas=2, espera_base=15
                )
            except Exception as e:
                msg_e = str(e)
                if any(c in msg_e for c in ["400", "INVALID_ARGUMENT", "403", "PERMISSION_DENIED"]):
                    ri, res, nota = _rj_extrair_chunk((idx, cp, offset, total_pg, n, cli))
                    return ri, res, (nota + " | " if nota else "") + "fallback inline"
                raise

        t_extr = time.time()
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
            futures = {ex.submit(_worker_rj, i): i for i in range(n)}
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, res, nota = future.result()
                    parciais[idx] = res
                    c = len(parciais)
                    log[-1] = _barra_progresso(c, n)
                    rounds_left = math.ceil((n - c) / 4)
                    est_rest = rounds_left * AVG_MIN_POR_CHUNK_RJ
                    log.append(
                        f"   Chunk {idx+1}/{n} extraido"
                        + (f" [{nota}]" if nota else "")
                        + f" | {c}/{n} prontos | ~{est_rest}min restantes"
                    )
                except Exception as e:
                    i_f = futures[future]
                    log.append(f"   Erro no chunk {i_f+1}: {e}")
                yield "\n".join(log), "", "", ""

        t_extr_s = int(time.time() - t_extr)
        log.append(f"   Extracao total: {t_extr_s//60}min{t_extr_s%60:02d}s")
        yield "\n".join(log), "", "", ""

        # Merge
        log.append("\nConsolidando textos extraidos...")
        yield "\n".join(log), "", "", ""
        lista = [parciais.get(i, "") for i in range(n)]
        texto_merged = _rj_merge_textos(lista)
        log.append(f"   {len(texto_merged):,} caracteres de informacao extraida")
        yield "\n".join(log), "", "", ""

        # Cache
        log.append(f"\nConfigurando context cache ({model_cons})...")
        yield "\n".join(log), "", "", ""
        cache = _rj_obter_cache(client1, model_cons)
        log[-1] = "Cache configurado." if cache else "Cache nao disponivel — usando prompt completo."
        yield "\n".join(log), "", "", ""

        # Secao A
        secoes = []
        log.append(f"\nGerando Secao A — Recuperacao Judicial ({model_cons})...")
        log.append(f"   Aguardando {model_cons}... (pode levar alguns minutos)")
        yield "\n".join(log), "", "", ""

        secao_a = _rj_consolidar_secao_a(client1, texto_merged, instrucoes, cache, model_cons, versao_resumida)
        log[-1] = "   Secao A gerada."
        secoes.append(secao_a)
        yield "\n".join(log), "", "", ""

        # Processos relacionados
        if pdf_relacionados:
            pdf_paths_rel = [f.name if hasattr(f, "name") else str(f) for f in pdf_relacionados]
            log.append(f"\nProcessando {len(pdf_paths_rel)} processo(s) relacionado(s)...")
            yield "\n".join(log), "", "", ""
            secao_rel, texto_bruto_rel, avisos_rel = "", "", []
            for kind, payload in _rj_processar_relacionados(
                pdf_paths_rel, client1, client2, instrucoes, cache, model_cons
            ):
                if kind == "progress":
                    log.append(payload)
                    yield "\n".join(log), "", "", ""
                else:
                    secao_rel, texto_bruto_rel, avisos_rel = payload
            if texto_bruto_rel:
                # Soma ao texto bruto da RJ principal — sem isso, dados de execucoes que so
                # aparecem nos relacionados (nao repetidos no PDF da RJ) nao chegavam aos
                # checklists/Excel de credores, que priorizam texto_merged como fonte.
                texto_merged = (texto_merged + "\n\n" + texto_bruto_rel) if texto_merged else texto_bruto_rel
            if secao_rel and "nenhum" not in secao_rel.lower():
                log.append("   Processos relacionados analisados.")
                secoes.append(secao_rel)
            else:
                log.append("   Nenhuma informacao relevante nos processos relacionados.")
            for aviso in avisos_rel:
                log.append(f"   {aviso}")
            yield "\n".join(log), "", "", ""

        relatorio = "\n\n".join(s for s in secoes if s)
        t_total = int(time.time() - t_inicio)
        log.append(f"\nAnalise concluida em {t_total//60}min{t_total%60:02d}s | {len(relatorio):,} chars")
        yield "\n".join(log), relatorio, relatorio, texto_merged

    except Exception as exc:
        log.append(f"\nErro: {exc}")
        yield "\n".join(log), f"Erro:\n{exc}", "", ""

    finally:
        for cp, _, _, orig in todos_chunks:
            if cp != orig:
                try: os.remove(cp)
                except: pass
        for cp in _temp_comprimidos:
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


def rj_gerar_checklist(relatorio: str, texto_bruto: str = ""):
    # O relatório (já consolidado, com referências corretas) é a fonte principal;
    # o texto bruto extraído do processo só complementa o que o relatório não cobrir.
    relatorio = (relatorio or "").strip()
    texto_bruto = (texto_bruto or "").strip()
    if not relatorio and not texto_bruto:
        return gr.update(value=None, visible=False), "Gere uma análise primeiro."
    try:
        k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k1)
        path = gerar_checklist_rj(relatorio, texto_bruto, client, MODEL_RAPIDO_RJ)
        return gr.update(value=path, visible=True), "✅ Checklist RJ gerado — clique no arquivo para baixar."
    except Exception as e:
        return gr.update(value=None, visible=False), f"❌ Erro: {e}"


def rj_gerar_checklist_creditos(relatorio: str, texto_bruto: str = "", *campos):
    # Prioriza o texto OCR completo (RJ + execuções); o relatório resumido omite campos
    fonte = texto_bruto.strip() if texto_bruto and texto_bruto.strip() else (relatorio or "").strip()
    if not fonte:
        return gr.update(value=None, visible=False), "Gere uma análise primeiro."
    # campos = [nome_1..nome_N, doc_1..doc_N] — pareia e ignora credores sem nome
    meia = len(campos) // 2
    nomes, docs = campos[:meia], campos[meia:]
    creditores = [(str(nomes[i]).strip(), str(docs[i]).strip())
                  for i in range(meia) if nomes[i] and str(nomes[i]).strip()]
    try:
        k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k1)
        resultado = gerar_checklist_creditos(fonte, client, MODEL_RAPIDO_RJ, creditores=creditores or None)
        paths = resultado if isinstance(resultado, list) else [resultado]
        if creditores:
            msg = f"✅ {len(paths)} checklist(s) de crédito gerado(s) — um arquivo por credor, clique para baixar."
        elif len(paths) > 1:
            msg = f"✅ {len(paths)} credor(es) exequente(s) identificado(s) automaticamente — um checklist por credor, clique para baixar."
        else:
            msg = "✅ Checklist de créditos gerado (crédito identificado automaticamente)."
        return gr.update(value=paths, visible=True), msg
    except Exception as e:
        return gr.update(value=None, visible=False), f"❌ Erro: {e}"


# ── Excel de Credores ─────────────────────────────────────────────────────────

_CLASSE_ORDER = [
    "I - Trabalhista",
    "II - Garantia Real",
    "III - Quirografário",
    "IV - ME/EPP",
    "Extraconcursal",
]

_CLASSE_FILL = {
    "I - Trabalhista":     "FFF2CC",
    "II - Garantia Real":  "D9E1F2",
    "III - Quirografário": "E2EFDA",
    "IV - ME/EPP":         "FCE4D6",
    "Extraconcursal":      "F2F2F2",
}


def _normalizar_classe(raw: str) -> str:
    c = (raw or "").upper()
    if re.search(r'\bCLASSE\s*I\b', c) or any(x in c for x in ("TRABALH", "ACIDENT")):
        return "I - Trabalhista"
    if re.search(r'\bCLASSE\s*II\b', c) or "GARANTIA REAL" in c:
        return "II - Garantia Real"
    if re.search(r'\bCLASSE\s*III\b', c) or "QUIROGRAF" in c:
        return "III - Quirografário"
    if re.search(r'\bCLASSE\s*IV\b', c) or any(x in c for x in ("ME/EPP", "MICROEMPRES", "PEQUENO")):
        return "IV - ME/EPP"
    if "EXTRACONCURSAL" in c:
        return "Extraconcursal"
    return raw.strip() if raw else "Não classificado"


def _extrair_credores_json(texto: str, client, model: str) -> dict:
    prompt = (
        "Você está analisando o texto extraído de um processo de Recuperação Judicial. "
        "Sua tarefa é montar o Quadro Geral de Credores (QGC) mais completo e preciso possível.\n\n"

        "═══ REGRA DE PRIORIDADE DE FONTE ═══\n"
        "1. PRIORIDADE MÁXIMA — Edital do Art. 18 (edital consolidado/definitivo, publicado APÓS a "
        "apreciação das divergências e impugnações pelo AJ e pelo juízo). Nomes comuns: 'Edital Art. 18', "
        "'Edital Consolidado', 'Edital de Habilitação Definitivo', 'Quadro Geral Consolidado'.\n"
        "   → Se encontrar, use-o como lista base.\n"
        "2. PRIORIDADE SECUNDÁRIA — Se NÃO houver Edital Art. 18: use o primeiro edital (Art. 7-A ou "
        "edital inicial do AJ) como lista base e COMPLEMENTE com:\n"
        "   a) Manifestação do AJ sobre divergências/impugnações (ajusta valores, classe ou garantias)\n"
        "   b) Decisões judiciais sobre habilitações/impugnações\n"
        "3. FALLBACK — Se não houver nenhum edital formal, extraia da lista de credores que aparecer "
        "no texto (QGC provisório, planilha, tabela, etc.)\n\n"

        "═══ PARA CADA CREDOR, EXTRAIA ═══\n"
        "- nome: nome/razão social completo\n"
        "- cpf_cnpj: apenas dígitos (sem pontuação), ou null\n"
        "- classe: 'I - Trabalhista' | 'II - Garantia Real' | 'III - Quirografário' | 'IV - ME/EPP' | 'Extraconcursal'\n"
        "- natureza: trabalhista, bancário, fornecedor, tributário, debenturista, CRI, CRA, CCB, etc.\n"
        "- instrumento: CCB nº X, Contrato nº Y, NP, duplicata, etc.\n"
        "- garantia_tipo: alienação fiduciária de imóvel, hipoteca, penhor de ações, fiança, etc.\n"
        "- bem_garantia: descrição do bem ou imóvel dado em garantia\n"
        "- valor_original: número sem formatação (reais), ou null\n"
        "- data_base: DD/MM/AAAA, ou null\n"
        "- valor_atualizado: valor habilitado/atualizado em reais, ou null\n"
        "- status: 'Habilitado' | 'Impugnado' | 'Divergente' | 'Extraconcursal' | 'Incluído' | 'Excluído'\n"
        "- processo_habilitacao: nº do incidente de habilitação ou impugnação, ou null\n"
        "- pagina_referencia: página(s) onde a informação aparece, ex: 'p. 234', 'p. 45-46'\n"
        "- conflitos: descreva divergências entre AJ e credor, ou entre credores, sobre valor, classe "
        "ou garantia (com referência de página). null se não houver.\n"
        "- questoes_controversas: impugnações pendentes, questões não decididas, pontos que ainda "
        "dependem de decisão judicial (com página). null se não houver.\n"
        "- observacoes: outras informações relevantes\n\n"

        "═══ REGRAS CRÍTICAS ═══\n"
        "- TABELAS ESCANEADAS: mesmo que a OCR seja imperfeita, extraia TODOS os campos visíveis. "
        "Não omita nenhuma linha da tabela, mesmo que incompleta. Informe 'dados parciais (pág. X)' "
        "em observacoes quando a leitura for incerta.\n"
        "- Não omita nenhum credor, mesmo que incompleto.\n"
        "- Se o valor no edital divergir da manifestação do AJ, registre AMBOS em conflitos.\n"
        "- Se houver impugnação pendente (credor ou recuperando impugnou o crédito de outro), "
        "registre em questoes_controversas.\n"
        "- Informações de garantia e instrumento podem vir de manifestações do AJ ou das impugnações — "
        "complemente o que estiver no edital.\n\n"

        "Retorne APENAS JSON válido (sem markdown) com esta estrutura:\n"
        "{\n"
        "  \"edital_utilizado\": \"Art. 18 (definitivo)\" | \"Primeiro edital + manifestação AJ\" | \"Primeiro edital\" | \"Lista informal\",\n"
        "  \"paginas_edital\": \"ex: p. 234-267\" ou null,\n"
        "  \"observacao_fontes\": \"explique quais seções/páginas foram usadas e se há lacunas\",\n"
        "  \"credores\": [...]\n"
        "}\n\n"
        f"TEXTO EXTRAÍDO:\n{texto[:150_000]}"
    )
    config = types.GenerateContentConfig(response_mime_type="application/json")

    def _fn():
        return client.models.generate_content(model=model, contents=[prompt], config=config).text

    raw = _retry(_fn)
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"```(?:json)?\n?", "", raw).strip().rstrip("`").strip()
    try:
        data = json.loads(raw)
        if isinstance(data, list):
            return {"edital_utilizado": "desconhecido", "paginas_edital": None,
                    "observacao_fontes": None, "credores": data}
        return data
    except json.JSONDecodeError:
        return {"edital_utilizado": "erro", "paginas_edital": None,
                "observacao_fontes": None, "credores": []}


def _gerar_excel_credores(resultado: dict) -> str:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    credores           = resultado.get("credores", [])
    edital_utilizado   = resultado.get("edital_utilizado", "")
    paginas_edital     = resultado.get("paginas_edital", "")
    observacao_fontes  = resultado.get("observacao_fontes", "")

    COR_HEADER    = "1F4E79"
    COR_SUB       = "BDD7EE"
    COR_TOTAL     = "2E75B6"
    COR_CONFLITO  = "FFD0D0"   # vermelho claro — conflito identificado
    COR_CONTROVER = "FFF3CD"   # amarelo claro — questão controversa pendente
    COR_META      = "EBF3FB"   # azul muito claro — linha de metadados

    thin  = Side(style="thin", color="D0D0D0")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)

    COLS = [
        ("Nome do Credor",                  48),
        ("CPF/CNPJ",                        18),
        ("Classe",                           22),
        ("Natureza do Crédito",              22),
        ("Instrumento / Lastro",             32),
        ("Garantia (tipo)",                  26),
        ("Bem dado em Garantia",             36),
        ("Valor Original (R$)",              18),
        ("Data Base",                        12),
        ("Valor Atualizado (R$)",            18),
        ("Status",                           16),
        ("Proc. Habilitação / Impugnação",   30),
        ("Página(s)",                        14),
        ("Conflitos (divergências)",         44),
        ("Questões Controversas",            44),
        ("% do Total",                       11),
        ("% da Classe",                      11),
        ("Observações",                      42),
    ]
    N = len(COLS)
    last_col = get_column_letter(N)

    wb = Workbook()
    ws = wb.active
    ws.title = "QGC"

    # ── Linha 1: título ────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value = "QUADRO GERAL DE CREDORES — RECUPERAÇÃO JUDICIAL"
    t.font      = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    t.fill      = PatternFill("solid", fgColor=COR_HEADER)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ── Linha 2: metadados de fonte ───────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    meta_txt = f"Edital base: {edital_utilizado}"
    if paginas_edital:
        meta_txt += f"  |  Páginas: {paginas_edital}"
    if observacao_fontes:
        meta_txt += f"  |  {observacao_fontes}"
    m = ws["A2"]
    m.value     = meta_txt
    m.font      = Font(name="Calibri", italic=True, size=9, color="1F4E79")
    m.fill      = PatternFill("solid", fgColor=COR_META)
    m.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 28

    # ── Linha 3: legenda de cores ─────────────────────────────────────────────
    ws.merge_cells(f"A3:{last_col}3")
    leg = ws["A3"]
    leg.value     = "Legenda:   Fundo vermelho = conflito identificado entre partes   |   Fundo amarelo = questão controversa pendente de decisão"
    leg.font      = Font(name="Calibri", italic=True, size=8, color="555555")
    leg.fill      = PatternFill("solid", fgColor="F8F8F8")
    leg.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[3].height = 18

    # ── Linha 4: cabeçalho de colunas ─────────────────────────────────────────
    HR = 4
    for ci, (h, w) in enumerate(COLS, 1):
        cell = ws.cell(row=HR, column=ci, value=h)
        cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        cell.fill      = PatternFill("solid", fgColor=COR_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = borda
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[HR].height = 34
    ws.freeze_panes = ws.cell(row=HR + 1, column=1)

    total_geral = sum(float(c.get("valor_atualizado") or 0) for c in credores)

    grupos: dict = defaultdict(list)
    for c in credores:
        grupos[_normalizar_classe(c.get("classe", ""))].append(c)

    def fv(v):
        try: return float(v) if v not in (None, "") else None
        except: return None

    cur = HR + 1
    classes_ord   = [cl for cl in _CLASSE_ORDER if cl in grupos]
    classes_extra = [cl for cl in grupos if cl not in _CLASSE_ORDER]

    for classe in classes_ord + classes_extra:
        lista        = grupos[classe]
        fill_cl      = PatternFill("solid", fgColor=_CLASSE_FILL.get(classe, "FFFFFF"))
        total_classe = sum(float(c.get("valor_atualizado") or 0) for c in lista)

        for credor in lista:
            va = fv(credor.get("valor_atualizado"))
            vo = fv(credor.get("valor_original"))
            pct_tot = (va / total_geral)  if (va and total_geral)  else None
            pct_cl  = (va / total_classe) if (va and total_classe) else None

            tem_conflito  = bool(credor.get("conflitos"))
            tem_controver = bool(credor.get("questoes_controversas"))

            if tem_conflito:
                fill_row = PatternFill("solid", fgColor=COR_CONFLITO)
            elif tem_controver:
                fill_row = PatternFill("solid", fgColor=COR_CONTROVER)
            else:
                fill_row = fill_cl

            row_vals = [
                credor.get("nome"),
                credor.get("cpf_cnpj"),
                classe,
                credor.get("natureza"),
                credor.get("instrumento"),
                credor.get("garantia_tipo"),
                credor.get("bem_garantia"),
                vo,
                credor.get("data_base"),
                va,
                credor.get("status"),
                credor.get("processo_habilitacao"),
                credor.get("pagina_referencia"),
                credor.get("conflitos"),
                credor.get("questoes_controversas"),
                pct_tot,
                pct_cl,
                credor.get("observacoes"),
            ]

            WRAP_COLS = {1, 5, 7, 13, 14, 15, 18}
            for ci, val in enumerate(row_vals, 1):
                cell = ws.cell(row=cur, column=ci, value=val)
                cell.fill      = fill_row
                cell.font      = Font(name="Calibri", size=10)
                cell.border    = borda
                cell.alignment = Alignment(vertical="center", wrap_text=(ci in WRAP_COLS))
                if ci in (8, 10):
                    cell.number_format = '#,##0.00'
                elif ci in (16, 17):
                    cell.number_format = '0.00%'
            cur += 1

        # Subtotal da classe
        sf = PatternFill("solid", fgColor=COR_SUB)
        sb = Font(name="Calibri", bold=True, size=10)
        for ci in range(1, N + 1):
            c = ws.cell(row=cur, column=ci)
            c.fill = sf; c.font = sb; c.border = borda
        ws.cell(row=cur, column=1).value = f"Subtotal — {classe}  ({len(lista)} credores)"
        c10 = ws.cell(row=cur, column=10, value=total_classe)
        c10.number_format = '#,##0.00'
        if total_geral:
            c16 = ws.cell(row=cur, column=16, value=total_classe / total_geral)
            c16.number_format = '0.00%'
        cur += 1

    # Total geral
    tf   = PatternFill("solid", fgColor=COR_TOTAL)
    tfnt = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
    for ci in range(1, N + 1):
        c = ws.cell(row=cur, column=ci)
        c.fill = tf; c.font = tfnt; c.border = borda
    ws.cell(row=cur, column=1).value = f"TOTAL GERAL  ({len(credores)} credores)"
    c10 = ws.cell(row=cur, column=10, value=total_geral)
    c10.number_format = '#,##0.00'
    ws.row_dimensions[cur].height = 22

    ws.auto_filter.ref = f"A{HR}:{last_col}{cur - 1}"

    out = os.path.join(tempfile.gettempdir(), "qgc_credores_rj.xlsx")
    wb.save(out)
    return out


def rj_gerar_excel_credores(relatorio: str, texto_bruto: str = ""):
    fonte = texto_bruto.strip() if texto_bruto and texto_bruto.strip() else relatorio.strip()
    if not fonte:
        return gr.update(value=None, visible=False), "Gere uma análise primeiro."
    try:
        k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k1)
        resultado = _extrair_credores_json(fonte, client, MODEL_CONSOLIDACAO_RJ)
        credores = resultado.get("credores", [])
        if not credores and texto_bruto and relatorio.strip() and fonte != relatorio.strip():
            resultado = _extrair_credores_json(relatorio, client, MODEL_CONSOLIDACAO_RJ)
            credores = resultado.get("credores", [])
        if not credores:
            return gr.update(value=None, visible=False), "Nenhum credor identificado. Verifique se o QGC consta no PDF."
        path = _gerar_excel_credores(resultado)
        edital = resultado.get("edital_utilizado", "")
        conflitos = sum(1 for c in credores if c.get("conflitos"))
        controver = sum(1 for c in credores if c.get("questoes_controversas"))
        msg = f"{len(credores)} credor(es) exportados"
        if edital:
            msg += f" | Edital base: {edital}"
        if conflitos:
            msg += f" | {conflitos} com conflito"
        if controver:
            msg += f" | {controver} com questão controversa"
        return gr.update(value=path, visible=True), msg
    except Exception as e:
        return gr.update(value=None, visible=False), f"Erro: {e}"



