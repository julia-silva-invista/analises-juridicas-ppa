# -*- coding: utf-8 -*-
"""Geração dos Checklists de Recuperação Judicial em Word (Invista PPA).

Dois documentos:
  1. Checklist de Recuperação Judicial       -> gerar_checklist_rj
  2. Análise de Créditos em Recuperação Judicial -> gerar_checklist_creditos

Os dados são extraídos preferencialmente do TEXTO OCR COMPLETO da análise
(não apenas do relatório resumido), pois o relatório omite campos do checklist.
"""

import json
import os
import re
import tempfile
import unicodedata
from datetime import date as _date

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from google.genai import types

from utils import _retry

# Reuso dos helpers de formatação do dossiê PPA (mesma paleta/fonte Invista)
from dossie_ppa import (
    _LARANJA, _CINZA, _BRANCO, _TXT, _TXT_MUTE,
    _write, _apply_font, _orange_header, _grid_table,
    _kv_table, _kv_label_table, _para, _sec_title, _sub_orange, _sub_gray,
    _note, _spacer, _montar_cabecalho_rodape,
)


# ══════════════════════════════════════════════════════════════════════════
# Helpers específicos de checklist
# ══════════════════════════════════════════════════════════════════════════

def _norm_cb(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"\s+", " ", s).strip()


def _cb(options, selected=""):
    """Marca ☑ APENAS uma opção — a que melhor casa com o valor extraído.

    Usa fronteira de palavra (\\b) para não marcar 'Favorável' quando o valor é
    'Desfavorável' (favoravel é substring de desfavoravel).
    """
    sel = _norm_cb(selected)
    if sel.startswith("nao consta"):
        sel = ""  # sentinel de "informação não encontrada" — nunca é uma opção marcável
    best_i, best_score = -1, 0
    if sel:
        for i, opt in enumerate(options):
            o = _norm_cb(opt)
            if not o:
                continue
            if sel == o:
                score = 1000
            elif re.search(r"\b" + re.escape(o) + r"\b", sel):
                score = 500 + len(o)      # a opção aparece inteira no valor
            elif re.search(r"\b" + re.escape(sel) + r"\b", o):
                score = 200 + len(sel)    # o valor aparece inteiro na opção
            else:
                score = 0
            if score > best_score:
                best_score, best_i = score, i
    return "   ".join(("☑ " if i == best_i and best_score > 0 else "☐ ") + opt
                      for i, opt in enumerate(options))


def _as_list(value) -> list:
    """Extração pode devolver tipo errado num campo-lista (PDF ruim/confuso).
    Nunca deixa isso quebrar a geração do documento."""
    return value if isinstance(value, list) else []


def _as_dict(value) -> dict:
    """Mesma defesa de _as_list, para campos que devem ser dict."""
    return value if isinstance(value, dict) else {}


def _titulo(doc, titulo, subtitulo="PPA Invista"):
    _montar_cabecalho_rodape(doc)
    _para(doc, titulo, bold=True, size=18, color=_LARANJA, before=2, after=0)
    _para(doc, subtitulo, bold=False, size=11, color=_TXT_MUTE, before=0, after=8)


def _rodape_conf(doc):
    _para(doc,
          "Documento confidencial. Uso interno exclusivo do Time PPA — Invista. "
          "Vedada a reprodução ou distribuição a terceiros sem autorização prévia.",
          size=8, color=_TXT_MUTE, italic=True, before=10, after=0,
          align=WD_ALIGN_PARAGRAPH.CENTER)


def _info_rj(doc, dados, hoje):
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Recuperação Judicial nº", dados.get("rj_numero", "")),
        ("Vara",                    dados.get("vara", "")),
        ("Data da Análise",         dados.get("data_analise") or hoje),
        ("Advogada Responsável",    ""),
    ])
    _spacer(doc)


def _salvar(doc, prefixo, dados):
    nome = re.sub(r"[^\w\s-]", "", dados.get("rj_numero", "") or "caso").strip().replace(" ", "_")[:40] or "caso"
    caminho = os.path.join(tempfile.gettempdir(), f"{prefixo}_{nome}.docx")
    doc.save(caminho)
    return caminho


def _base_doc():
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9.5)
    return doc


def _extrair(prompt_base, fonte, client, model):
    prompt = prompt_base + fonte[:900_000]
    try:
        config = types.GenerateContentConfig(response_mime_type="application/json")

        def _fn():
            return client.models.generate_content(
                model=model,
                contents=[types.Content(role="user", parts=[types.Part(text=prompt)])],
                config=config,
            ).text

        raw = _retry(_fn, tentativas=3, espera_base=10)
        raw = re.sub(r"^```[a-z]*\n?", "", raw.strip(), flags=re.IGNORECASE)
        raw = re.sub(r"\n?```$", "", raw.strip())
        dados = json.loads(raw)
        return dados if isinstance(dados, dict) else {}
    except Exception:
        return {}


# ══════════════════════════════════════════════════════════════════════════
# 1. CHECKLIST DE RECUPERAÇÃO JUDICIAL
# ══════════════════════════════════════════════════════════════════════════

# Documentos do item 6 (chave, rótulo, opções de status). A chave é usada tanto no
# checklist quanto no recorte do PDF (rj._localizar_e_recortar_docs).
_DOCS_ITEM6 = [
    ("peticao_inicial",       "Petição Inicial da RJ",                              ["Salvo", "Pendente"]),
    ("quadro_ativos",         "Quadro de Ativos dos Requerentes",                   ["Anexado", "Não existente/Segredo de Justiça"]),
    ("pericia_previa",        "Relatório de Perícia Prévia",                        ["Anexado", "Não existente"]),
    ("laudo_imoveis",         "Laudo de Imóveis Essenciais",                        ["Anexado", "Não existente"]),
    ("ultimo_rma",            "Último RMA",                                         ["Anexado", "Não existente"]),
    ("qgc_recuperando",       "QGC do(s) recuperando(s)",                           ["Anexado", "Não existente"]),
    ("qgc_aj",                "QGC do AJ",                                          ["Anexado", "Não existente"]),
    ("relatorio_divergencia", "Relatório de Divergência",                           ["Anexado", "Não existente"]),
    ("prj_aditivos",          "PRJ e Aditivos",                                     ["Anexado", "Não existente"]),
    ("atas_agc",              "Atas, laudo de credenciamento e de votação da AGC",  ["Anexado", "Não existente"]),
]

_PROMPT_RJ = """Você está analisando o texto COMPLETO extraído de um processo de Recuperação Judicial (RJ).
Analise TODO o texto (não só um resumo) para preencher o Checklist de RJ no formato JSON abaixo.
NÃO invente.

═══ REGRA — SISTEMA DO TRIBUNAL ═══
Primeiro identifique, pelo cabeçalho/rodapé/numeração do PDF, qual sistema processual gerou o
documento, e use SEMPRE o rótulo correspondente nas referências:
  - PJe                → "Mov. <nº>"
  - eSAJ / físico       → "fls. <nº>"
  - Eproc               → "Evento <nº>"
  - Projudi / outro     → "ID <nº>"
Use o MESMO rótulo em todo o documento (não misture "fls." com "Mov." sem necessidade).

═══ REGRA — REFERÊNCIA OBRIGATÓRIA DA FONTE (SEM EXCEÇÃO) ═══
TODA informação que você preencher — cada campo, cada linha de cada tabela, cada valor, data, nome,
matrícula, condição do PRJ, classe do QGC, situação da AGC, endividamento, etc. — deve trazer ao final,
entre parênteses, a referência exata de onde foi extraída (ex.: "(Mov. 340)", "(fls. 88)",
"(Evento 12)"). Isso vale para TODOS os campos do JSON abaixo, mesmo os que não têm um exemplo de
referência escrito explicitamente — o exemplo "(fls.)"/"(Mov./ID/fls./Evento)" que aparece no molde
abaixo é só uma INSTRUÇÃO DE FORMATO PARA VOCÊ, não é texto para copiar. Você deve SEMPRE substituí-lo
pela referência real e específica daquele dado.
NUNCA devolva a anotação de formato vazia ou literal (nunca escreva "(fls.)" sozinho, sem número, nem
"(Mov./ID/fls./Evento)" literalmente). Se um campo tiver informação mas você não conseguir localizar a
página/movimentação exata dele, ainda assim preencha o valor e finalize com "(referência não localizada)"
em vez de inventar um número.

═══ REGRA — CAMPO NÃO ENCONTRADO ═══
Se um dado factual não constar em NENHUMA parte do processo, preencha com "Não consta" (sem parênteses).
Em campos de escolha, responda com a opção EXATA e ÚNICA (ex: "Deferido (Mov. 12)"). NUNCA marque mais de
uma opção; se não houver informação, deixe "" (nenhuma opção marcada).

Responda SOMENTE com o JSON.

{
  "rj_numero": "", "vara": "", "data_analise": "",
  "requerentes": "Nome · CPF/CNPJ ... (referência)",
  "advogados_requerentes": "(referência)",
  "administrador_judicial": "Nome (referência)",
  "data_pedido": "DD/MM/AAAA (referência)", "data_deferimento": "DD/MM/AAAA (referência)",
  "consolidacao_substancial": "Deferido/Indeferido/... (referência)",
  "periodo_blindagem": "Ativo/Inativo (referência)",
  "previsao_encerramento_stay": "DD/MM/AAAA (referência)",
  "stay_prorrogavel": "Sim (referência) | Não (referência) | ''",
  "recursos_relevantes": [{"recurso": "(referência)", "status": "(referência)"}],
  "imoveis_requerentes": [{"matricula": "nº (referência)", "cartorio": "(referência)", "descricao": "(referência)", "proprietario": "(referência)"}],
  "imoveis_essenciais": [{"matricula": "nº (referência)", "cartorio": "(referência)", "descricao": "(referência)", "proprietario": "(referência)"}],
  "prj_classe_ii": {"desagio": "(referência)", "carencia": "(referência)", "parcelas": "(referência)", "juros": "(referência)", "correcao": "(referência)"},
  "prj_classe_iii": {"desagio": "(referência)", "carencia": "(referência)", "parcelas": "(referência)", "juros": "(referência)", "correcao": "(referência)"},
  "qgc": {"classe_i": "R$ (referência)", "classe_ii": "R$ (referência)", "classe_iii": "R$ (referência)", "classe_iv": "R$ (referência)", "total": "R$ (referência)"},
  "agc_situacao": "opção (referência)",
  "agc_1a": "DD/MM/AAAA (referência)", "agc_2a": "DD/MM/AAAA (referência)", "agc_continuacao": "(referência)",
  "recuperandos": [{"nome": "(referência)", "ecac": "R$ (referência)", "divida_ativa": "R$ (referência)"}],
  "endividamento_fiscal_total": "R$ (referência)",
  "documentos_salvos": {
    "peticao_inicial":       {"status": "Salvo | Pendente", "folhas": "referência exata onde consta"},
    "quadro_ativos":         {"status": "Anexado | Não existente", "folhas": ""},
    "pericia_previa":        {"status": "Anexado | Não existente", "folhas": ""},
    "laudo_imoveis":         {"status": "Anexado | Não existente", "folhas": ""},
    "ultimo_rma":            {"status": "Anexado | Não existente", "folhas": ""},
    "qgc_recuperando":       {"status": "Anexado | Não existente", "folhas": ""},
    "qgc_aj":                {"status": "Anexado | Não existente", "folhas": ""},
    "relatorio_divergencia": {"status": "Anexado | Não existente", "folhas": ""},
    "prj_aditivos":          {"status": "Anexado | Não existente", "folhas": ""},
    "atas_agc":              {"status": "Anexado | Não existente", "folhas": ""}
  }
}

Para documentos_salvos: informe, para cada documento, se ele está anexado/salvo no processo e a
referência exata (fls./Mov./ID/Evento) onde ele aparece — sem os colchetes do molde de preenchimento
manual ("[caso haja, indicar páginas]"): esse texto é só a instrução do molde para humanos e NUNCA deve
aparecer no seu JSON; escreva o número/referência real, ou deixe "folhas" vazio se o documento não
existir no processo (status "Não existente").

TEXTO:
"""


def _build_checklist_rj(dados: dict) -> str:
    doc = _base_doc()
    hoje = _date.today().strftime("%d/%m/%Y")
    _titulo(doc, "Checklist de Recuperação Judicial")
    _info_rj(doc, dados, hoje)

    # ── 1. DADOS GERAIS DO CASO ──────────────────────────────────────────
    _sec_title(doc, "1. DADOS GERAIS DO CASO")
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Requerentes (Nome e CPF/CNPJ)", dados.get("requerentes", "")),
        ("Advogados dos Requerentes",     dados.get("advogados_requerentes", "")),
        ("Administrador Judicial Nomeado", dados.get("administrador_judicial", "")),
        ("Data do Pedido",                dados.get("data_pedido", "")),
        ("Data do Deferimento",           dados.get("data_deferimento", "")),
        ("Consolidação Substancial",      _cb(["Solicitado, ainda sem decisão", "Deferido", "Indeferido", "Não se aplica"],
                                              dados.get("consolidacao_substancial", ""))),
        ("Período de Blindagem",          _cb(["Ativo", "Inativo"], dados.get("periodo_blindagem", ""))),
        ("Previsão de Encerramento do Stay",
         (dados.get("previsao_encerramento_stay", "") or "____") + "    " +
         _cb(["Prorrogável por mais 180 dias"], dados.get("stay_prorrogavel", ""))),
    ])
    _spacer(doc, pts=2)

    # Recursos Relevantes
    _sub_gray(doc, "Recursos Relevantes")
    _grid_table(
        doc, ["RECURSO RELEVANTE", "STATUS"],
        [[_as_dict(r).get("recurso", ""), _as_dict(r).get("status", "")] for r in _as_list(dados.get("recursos_relevantes"))],
        [10.9, 6.0], min_rows=1,
    )
    _spacer(doc, pts=2)

    # Imóveis dos Requerentes / Essenciais
    cols_imo = ["Nº da Matrícula", "Cartório", "Descrição", "Proprietário"]
    ws_imo = [3.0, 3.5, 6.4, 4.0]
    _sub_gray(doc, "Imóveis dos Requerentes")
    _grid_table(doc, cols_imo,
                [[_as_dict(i).get("matricula", ""), _as_dict(i).get("cartorio", ""),
                  _as_dict(i).get("descricao", ""), _as_dict(i).get("proprietario", "")]
                 for i in _as_list(dados.get("imoveis_requerentes"))],
                ws_imo, min_rows=2)
    _spacer(doc, pts=2)
    _sub_gray(doc, "Imóveis Essenciais")
    _grid_table(doc, cols_imo,
                [[_as_dict(i).get("matricula", ""), _as_dict(i).get("cartorio", ""),
                  _as_dict(i).get("descricao", ""), _as_dict(i).get("proprietario", "")]
                 for i in _as_list(dados.get("imoveis_essenciais"))],
                ws_imo, min_rows=2)
    _spacer(doc)

    # ── 2. PLANO DE RECUPERAÇÃO JUDICIAL (PRJ) ───────────────────────────
    _sec_title(doc, "2. PLANO DE RECUPERAÇÃO JUDICIAL (PRJ)")
    cols_cond = ["Deságio", "Carência", "Parcelas", "Juros", "Correção"]
    ws_cond = [3.4, 3.4, 3.4, 3.35, 3.35]
    for classe, key in [("Condições da Classe II", "prj_classe_ii"),
                        ("Condições da Classe III", "prj_classe_iii")]:
        _sub_gray(doc, classe)
        c = _as_dict(dados.get(key))
        _grid_table(doc, cols_cond,
                    [[c.get("desagio", ""), c.get("carencia", ""), c.get("parcelas", ""),
                      c.get("juros", ""), c.get("correcao", "")]],
                    ws_cond, min_rows=1)
        _spacer(doc, pts=2)
    _spacer(doc)

    # ── 3. QUADRO GERAL DE CREDORES (QGC) ────────────────────────────────
    _sec_title(doc, "3. QUADRO GERAL DE CREDORES (QGC):")
    qgc = _as_dict(dados.get("qgc"))
    _kv_table(doc, ["CLASSE", "VALOR"], [
        ("Classe I",   qgc.get("classe_i", "R$")),
        ("Classe II",  qgc.get("classe_ii", "R$")),
        ("Classe III", qgc.get("classe_iii", "R$")),
        ("Classe IV",  qgc.get("classe_iv", "R$")),
        ("Total dos créditos arrolados", qgc.get("total", "R$")),
    ])
    _spacer(doc)

    # ── 4. ASSEMBLEIA GERAL DE CREDORES (AGC) ────────────────────────────
    _sec_title(doc, "4. ASSEMBLEIA GERAL DE CREDORES (AGC)")
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Situação da AGC", _cb(["Sem datas designadas", "Convocada", "1ª convocação sem quórum",
                                 "Período de suspensão", "Plano aprovado / rejeitado"],
                                dados.get("agc_situacao", ""))),
        ("1ª Convocação",              dados.get("agc_1a", "")),
        ("2ª Convocação",              dados.get("agc_2a", "")),
        ("Continuação da 2ª Convocação", dados.get("agc_continuacao", "")),
    ])
    _spacer(doc)

    # ── 5. ENDIVIDAMENTO GERAL ───────────────────────────────────────────
    _sec_title(doc, "5. ENDIVIDAMENTO GERAL")
    recuperandos = _as_list(dados.get("recuperandos")) or [{}]
    for i, rec in enumerate(recuperandos, 1):
        rec = _as_dict(rec)
        nome = rec.get("nome", "") or "Não consta"
        _sub_gray(doc, f"RECUPERANDO {i} — {nome}")
        _kv_table(doc, ["ENDIVIDAMENTO FISCAL", "VALOR / STATUS"], [
            ("Endividamento Fiscal — e-CAC",
             (rec.get("ecac", "") or "R$") + "    " + _cb(["CND na pasta", "Não foi possível emitir"])),
            ("Endividamento Fiscal — Dívida Ativa",
             (rec.get("divida_ativa", "") or "R$") + "    " + _cb(["CND na pasta", "Não foi possível emitir"])),
        ])
        _spacer(doc, pts=2)
    _sub_gray(doc, "TOTAL CONSOLIDADO DO GRUPO")
    _kv_table(doc, ["CAMPO", "VALOR"], [
        ("Endividamento Fiscal Total", dados.get("endividamento_fiscal_total", "R$")),
    ])
    _note(doc, "Soma de e-CAC + Dívida Ativa de todos os recuperandos.")
    _spacer(doc)

    # ── 6. CHECKLIST DOS DOCUMENTOS SALVOS ───────────────────────────────
    _sec_title(doc, "6. CHECKLIST DOS DOCUMENTOS SALVOS")
    docs_salvos = _as_dict(dados.get("documentos_salvos"))
    rows6 = []
    for key, label, opts in _DOCS_ITEM6:
        info = _as_dict(docs_salvos.get(key))
        rows6.append((label, _cb(opts, info.get("status", "")), info.get("folhas", "")))
    _grid_table(doc, ["DOCUMENTO", "STATUS", "REFERÊNCIA"], rows6, [8.0, 5.9, 3.0])

    _rodape_conf(doc)
    return _salvar(doc, "Checklist_RJ", dados)


def gerar_checklist_rj(fonte: str, client, model: str) -> str:
    dados = _extrair(_PROMPT_RJ, fonte, client, model)
    return _build_checklist_rj(dados)


# ══════════════════════════════════════════════════════════════════════════
# 2. ANÁLISE DE CRÉDITOS EM RECUPERAÇÃO JUDICIAL
# ══════════════════════════════════════════════════════════════════════════

_PROMPT_CRED = """\
Você está analisando o texto COMPLETO extraído de um processo de Recuperação Judicial e das
execuções relacionadas a um crédito. Extraia os dados do CRÉDITO no formato JSON abaixo.
Analise TODO o texto (não só um resumo) para preencher cédula, emitente, avalista, garantias, etc.
NÃO invente.

═══ REGRA — REFERÊNCIA OBRIGATÓRIA DA FONTE ═══
Para CADA informação preenchida, inclua ao final, entre parênteses, ONDE foi extraída — a fls. do PDF
e o parâmetro do tribunal (Mov./ID/fls./Evento). Ex.: "CCB nº 123 (fls. 45)", "R$ 8.000.000 (fls. 88)".

═══ REGRA — CAMPO NÃO ENCONTRADO ═══
Se um dado factual (cédula, emitente, avalista, matrícula, valor, data, etc.) NÃO constar em NENHUMA
parte do processo, preencha com "Não consta". Em campos de escolha (Sim/Não, Favorável/...), se não
houver informação, deixe "" (nenhuma opção marcada). NUNCA marque mais de uma opção.

Em campos de escolha, responda com a opção EXATA e ÚNICA (ex: "Favorável", "Desfavorável", "Sim").
Responda SOMENTE com o JSON.

{
  "rj_numero": "", "vara": "", "data_analise": "",
  "credor": "Nome · CNPJ", "advogados": "",
  "classe_ii_valor": "R$", "classe_ii_garantias": "", "classe_ii_repr": "",
  "classe_iii_valor": "R$", "classe_iii_repr": "",
  "extraconcursal_valor": "R$", "extraconcursal_garantias": "",
  "recurso_credor_alvo": "Sim | Não",
  "impugnacoes": [
    {"numero": "", "polo_ativo": "", "finalidade": "", "lastro": "",
     "manifestacao_aj": "Favorável | Desfavorável | Pendente",
     "manifestacao_mp": "Favorável | Desfavorável | Pendente",
     "sentenca": "Favorável | Desfavorável | Pendente", "status": ""}
  ],
  "lastros": [
    {"cedula": "", "emitentes": "", "avalistas": "", "coobrigados_rj": "Sim | Não",
     "emissao": "", "garantia": "", "percentual": "", "valor_arrolado": "R$",
     "classe": "", "acoes": "", "obs_extraconcursal": ""}
  ],
  "garantias": [
    {"matricula": "", "comarca": "", "proprietario": "", "descricao": "", "onus": "",
     "avaliacao": "R$", "proprietarios_rj": "Sim | Não (qual?)"}
  ],
  "execucoes": [
    {"numero": "", "polo_ativo": "", "polo_passivo": "", "distribuicao": "", "lastro_fls": "",
     "garantia": "", "valor_causa": "R$", "honorarios": "", "sucumbencia": "",
     "constricao": "", "status": "", "cumprimento_autonomo": ""}
  ],
  "recursos": [{"numero": "", "polo_ativo": "", "finalidade": "", "status": ""}]
}

TEXTO:
"""


def _creditor_sections(doc, dados: dict):
    # ── 1. RESUMO DO CRÉDITO ─────────────────────────────────────────────
    _sec_title(doc, "1. RESUMO DO CRÉDITO")
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Credor (Nome e CNPJ)",        dados.get("credor", "")),
        ("Advogado(s)",                 dados.get("advogados", "")),
        ("Classe II",                   dados.get("classe_ii_valor", "R$")),
        ("Garantia(s) — Classe II",     dados.get("classe_ii_garantias", "")),
        ("Representatividade (%) — Classe II", dados.get("classe_ii_repr", "")),
        ("Classe III",                  dados.get("classe_iii_valor", "R$")),
        ("Representatividade (%) — Classe III", dados.get("classe_iii_repr", "")),
        ("Extraconcursal",              dados.get("extraconcursal_valor", "R$")),
        ("Garantia(s) — Extraconcursal", dados.get("extraconcursal_garantias", "")),
        ("Recurso(s) do credor-alvo?",  _cb(["Sim", "Não"], dados.get("recurso_credor_alvo", ""))),
    ])
    _spacer(doc)

    # ── 2. IMPUGNAÇÃO DE CRÉDITO ─────────────────────────────────────────
    _sec_title(doc, "2. IMPUGNAÇÃO DE CRÉDITO")
    impugnacoes = dados.get("impugnacoes") or [{}]
    manif = ["Favorável", "Desfavorável", "Pendente"]
    for i, imp in enumerate(impugnacoes, 1):
        _sub_gray(doc, f"Impugnação de Crédito nº {i}")
        _kv_label_table(doc, [
            ("Nº do Processo",              imp.get("numero", "")),
            ("Polo Ativo",                  imp.get("polo_ativo", "")),
            ("Finalidade",                  imp.get("finalidade", "")),
            ("Lastro em discussão",         imp.get("lastro", "")),
            ("Houve manifestação do AJ?",   _cb(manif, imp.get("manifestacao_aj", ""))),
            ("Houve manifestação do MP?",   _cb(manif, imp.get("manifestacao_mp", ""))),
            ("Sentença",                    _cb(manif, imp.get("sentenca", ""))),
            ("Status",                      imp.get("status", "")),
        ])
        _spacer(doc, pts=2)

    # ── 3. LASTROS ───────────────────────────────────────────────────────
    _sec_title(doc, "3. LASTROS")
    lastros = dados.get("lastros") or [{}]
    for i, la in enumerate(lastros, 1):
        _sub_gray(doc, f"Lastro nº {i}")
        _kv_label_table(doc, [
            ("Cédula",                        la.get("cedula", "")),
            ("Emitente(s)",                   la.get("emitentes", "")),
            ("Avalista(s)/Fiador(es)",        la.get("avalistas", "")),
            ("Todos os coobrigados estão em RJ", _cb(["Sim", "Não"], la.get("coobrigados_rj", ""))),
            ("Emissão",                       la.get("emissao", "")),
            ("Garantia",                      la.get("garantia", "")),
            ("Percentual/Limite garantido",   la.get("percentual", "")),
            ("Valor Arrolado",                la.get("valor_arrolado", "R$")),
            ("Classe",                        la.get("classe", "")),
            ("Ações relacionadas",            la.get("acoes", "")),
            ("Observação - Extraconcursal",   la.get("obs_extraconcursal", "")),
        ])
        _spacer(doc, pts=2)

    # ── 4. GARANTIAS ─────────────────────────────────────────────────────
    _sec_title(doc, "4. GARANTIAS")
    garantias = dados.get("garantias") or [{}]
    for i, ga in enumerate(garantias, 1):
        _sub_gray(doc, f"Garantia nº {i}")
        _kv_label_table(doc, [
            ("Matrícula nº",   ga.get("matricula", "")),
            ("Comarca",        ga.get("comarca", "")),
            ("Proprietário",   ga.get("proprietario", "")),
            ("Descrição",      ga.get("descricao", "")),
            ("Ônus",           ga.get("onus", "")),
            ("Avaliação",      ga.get("avaliacao", "R$")),
            ("Todos os proprietários estão em RJ? Se não, qual?",
             _cb(["Sim", "Não"], ga.get("proprietarios_rj", "")) +
             (("  " + ga.get("proprietarios_rj", "")) if ga.get("proprietarios_rj", "") not in ("", "Sim", "Não") else "")),
        ])
        _spacer(doc, pts=2)

    # ── 5. AÇÕES JUDICIAIS ───────────────────────────────────────────────
    _sec_title(doc, "5. AÇÕES JUDICIAIS")
    execucoes = dados.get("execucoes") or [{}]
    for i, ex in enumerate(execucoes, 1):
        _sub_gray(doc, f"Execução de Título Extrajudicial nº {i}")
        _kv_label_table(doc, [
            ("Nº do Processo",           ex.get("numero", "")),
            ("Polo Ativo",               ex.get("polo_ativo", "")),
            ("Polo Passivo",             ex.get("polo_passivo", "")),
            ("Distribuição",             ex.get("distribuicao", "")),
            ("Lastro (fls.)",            ex.get("lastro_fls", "")),
            ("Garantia",                 ex.get("garantia", "")),
            ("Valor da Causa — R$ (fls.)", ex.get("valor_causa", "")),
            ("Honorários - credor",      ex.get("honorarios", "")),
            ("Sucumbência",              ex.get("sucumbencia", "")),
            ("Constrição",               ex.get("constricao", "")),
            ("Status",                   ex.get("status", "")),
            ("Existe cumprimento autônomo de honorários (CEF/BB)", ex.get("cumprimento_autonomo", "")),
        ])
        _spacer(doc, pts=2)

    recursos = dados.get("recursos") or []
    for i, rec in enumerate(recursos, 1):
        _sub_gray(doc, f"Recurso nº {i}")
        _kv_label_table(doc, [
            ("Nº do Processo",     rec.get("numero", "")),
            ("Polo ativo",         rec.get("polo_ativo", "")),
            ("Finalidade/Matéria", rec.get("finalidade", "")),
            ("Status Atual",       rec.get("status", "")),
        ])
        _spacer(doc, pts=2)


def _build_checklist_creditos(dados: dict) -> str:
    doc = _base_doc()
    hoje = _date.today().strftime("%d/%m/%Y")
    _titulo(doc, "Análise de Créditos em Recuperação Judicial")
    _info_rj(doc, dados, hoje)
    _creditor_sections(doc, dados)
    _rodape_conf(doc)
    return _salvar(doc, "Analise_Creditos_RJ", dados)


def _build_checklist_creditos_multi(rj_info: dict, lista: list) -> str:
    """Um único documento com um checklist por credor (quebra de página entre eles)."""
    doc = _base_doc()
    hoje = _date.today().strftime("%d/%m/%Y")
    _titulo(doc, "Análise de Créditos em Recuperação Judicial")
    _info_rj(doc, rj_info, hoje)
    for i, dados in enumerate(lista):
        if i > 0:
            doc.add_page_break()
        nome = dados.get("credor", "") or f"Credor {i+1}"
        _para(doc, f"▸ CRÉDITO {i+1} — {nome}", bold=True, size=12, color=_LARANJA, before=2, after=4)
        _creditor_sections(doc, dados)
    _rodape_conf(doc)
    return _salvar(doc, "Analise_Creditos_RJ", rj_info or (lista[0] if lista else {}))


def _extrair_cred_scoped(fonte, nome, doc, client, model) -> dict:
    foco = (
        "═══ FOCO OBRIGATÓRIO ═══\n"
        f'Extraia os dados APENAS do credor "{nome}"'
        + (f" (CPF/CNPJ: {doc})" if doc else "")
        + ". Ignore todos os outros credores do processo.\n\n"
    )
    return _extrair(foco + _PROMPT_CRED, fonte, client, model)


def gerar_checklist_creditos(fonte: str, client, model: str, creditores=None) -> str:
    # Sem credores informados: extrai o crédito que a IA encontrar (comportamento padrão)
    if not creditores:
        dados = _extrair(_PROMPT_CRED, fonte, client, model)
        return _build_checklist_creditos(dados)
    # Com credores: um checklist escopado por credor, tudo em um documento
    lista = []
    rj_info = {}
    for nome, doc_id in creditores:
        dados = _extrair_cred_scoped(fonte, nome, doc_id, client, model)
        if not (dados.get("credor") or "").strip():
            dados["credor"] = nome + (f" · {doc_id}" if doc_id else "")
        lista.append(dados)
        if not rj_info and (dados.get("rj_numero") or dados.get("vara")):
            rj_info = {"rj_numero": dados.get("rj_numero", ""), "vara": dados.get("vara", ""),
                       "data_analise": dados.get("data_analise", "")}
    return _build_checklist_creditos_multi(rj_info, lista)
