# -*- coding: utf-8 -*-
"""Geração dos Checklists de Recuperação Judicial em Word (Invista PPA).

Dois documentos:
  1. Checklist de Recuperação Judicial       -> gerar_checklist_rj
  2. Análise de Créditos em Recuperação Judicial -> gerar_checklist_creditos

O Checklist de Recuperação Judicial usa o RELATÓRIO já consolidado (revisado,
deduplicado, com referências processuais corretas) como fonte principal, e o
texto bruto extraído do processo só como complemento do que o relatório não
cobrir. A Análise de Créditos ainda usa só o texto bruto/relatório resumido
(fonte única) — fica para uma próxima rodada.
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
    _spacer, _montar_cabecalho_rodape,
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
    linha = "   ".join(("☑ " if i == best_i and best_score > 0 else "☐ ") + opt
                       for i, opt in enumerate(options))
    # Preserva a referência (Mov./fls./ID/Evento) que o modelo anexou ao valor
    # extraído — sem isso, campos de escolha perdiam a referência ao virar
    # só símbolos de checkbox.
    ref = re.search(r"\([^()]*\)\s*$", str(selected or ""))
    if ref and best_score > 0:
        linha += "  " + ref.group(0)
    return linha


def _as_list(value) -> list:
    """Extração pode devolver tipo errado num campo-lista (PDF ruim/confuso).
    Nunca deixa isso quebrar a geração do documento."""
    return value if isinstance(value, list) else []


def _as_dict(value) -> dict:
    """Mesma defesa de _as_list, para campos que devem ser dict."""
    return value if isinstance(value, dict) else {}


def _as_str(value) -> str:
    """Mesma defesa de _as_list/_as_dict, para campos que devem ser string
    (ex.: extração devolve um dict/lista onde se esperava texto simples)."""
    return value if isinstance(value, str) else ""


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


def _salvar(doc, prefixo, dados, sufixo: str = ""):
    nome = re.sub(r"[^\w\s-]", "", dados.get("rj_numero", "") or "caso").strip().replace(" ", "_")[:40] or "caso"
    if sufixo:
        suf = re.sub(r"[^\w\s-]", "", sufixo).strip().replace(" ", "_")[:30]
        if suf:
            nome = f"{nome}_{suf}"
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

_PROMPT_RJ = """Você vai preencher o Checklist de Recuperação Judicial (RJ) no formato JSON abaixo, a
partir de duas fontes de texto que serão fornecidas depois deste prompt.

═══ REGRA — PRIORIDADE DE FONTE ═══
A primeira fonte é o RELATÓRIO CONSOLIDADO — um relatório jurídico já sintetizado, revisado e
deduplicado sobre este mesmo processo, com as referências processuais corretas (Mov./ID/fls./Evento)
já embutidas no texto. Use o relatório como fonte PRINCIPAL para todo campo — é a fonte mais confiável
e mais completa. A segunda fonte é o TEXTO BRUTO EXTRAÍDO — a extração página a página do processo,
mais volumosa e não deduplicada. Consulte o texto bruto APENAS para complementar informação que o
relatório não cubra, ou para achar uma referência mais específica de algo que o relatório mencione sem
citar página exata. Nunca contrarie o relatório com base em uma leitura isolada do texto bruto — se os
dois divergirem, prefira o relatório.

═══ REGRA — SISTEMA DO TRIBUNAL ═══
Identifique, pelo cabeçalho/rodapé/numeração do processo, qual sistema processual gerou o
documento, e use SEMPRE o rótulo correspondente nas referências:
  - PJe                → "Mov. <nº>"
  - eSAJ / físico       → "fls. <nº>"
  - Eproc               → "Evento <nº>"
  - Projudi / outro     → "ID <nº>"
Use o MESMO rótulo em todo o documento (não misture "fls." com "Mov." sem necessidade). O texto bruto
pode conter marcadores internos do tipo "PARTE i/N" ou "páginas X-Y" — são só limites técnicos de
divisão do arquivo pelo nosso sistema, NUNCA uma referência processual real. Nunca cite "Parte N" como
se fosse referência oficial. Se não houver ID/Mov./Evento disponível e só houver número de página, cite
só a página (ex.: "pág. 340"), tratando o conjunto como um documento contínuo único — a menos que fique
evidente que são processos ou documentos efetivamente distintos (números de processo diferentes,
cabeçalhos de arquivo diferentes), caso em que identifique de qual documento veio.

═══ REGRA — REFERÊNCIA OBRIGATÓRIA, UMA VEZ POR LINHA ═══
Todo campo com conteúdo deve trazer ao final, entre parênteses, a referência de onde foi extraído
(ex.: "(Mov. 340)", "(fls. 88)", "(Evento 12)", "(pág. 340)") — EXCETO "rj_numero" e "vara", que vão
sem nenhuma referência (só o valor). O exemplo "(referência)" que aparece no molde abaixo é só uma
INSTRUÇÃO DE FORMATO PARA VOCÊ, não é texto para copiar — substitua sempre pela referência real.
NUNCA devolva a anotação de formato vazia ou literal (nunca escreva "(referência)" ou
"(Mov./ID/fls./Evento)" literalmente).
Em uma linha de tabela com várias colunas (ex.: um imóvel com matrícula/cartório/descrição/proprietário,
ou um recurso com nome/status), coloque a referência SÓ no primeiro campo da linha (ex.: só em
"matricula") — os demais campos da MESMA linha trazem só o conteúdo, sem repetir a referência.

═══ REGRA — CAMPO NÃO ENCONTRADO ═══
Se um dado factual não constar em NENHUMA das duas fontes, preencha com "Não consta" — SEM parênteses
de referência nenhuma (nem vazia, nem "referência não localizada": esse fallback só vale quando HÁ
conteúdo mas a referência exata não foi localizada, nunca quando o campo inteiro é "Não consta").
Em campos de escolha, responda com a opção EXATA e ÚNICA (ex: "Deferido (Mov. 12)"). NUNCA marque mais
de uma opção; se não houver informação, deixe "" (nenhuma opção marcada).

═══ REGRA — AGC SEM DATAS ═══
Para "agc_situacao" especificamente, se não houver NENHUMA informação de AGC no processo, não escreva
"Não consta" — escreva "Sem datas designadas" (é a opção correta pra "AGC ainda não convocada").

═══ REGRA — IMÓVEIS ESSENCIAIS ═══
"Cartório" é sempre o CRI (Cartório de Registro de Imóveis) — use a mesma terminologia do relatório
("X CRI de Comarca/UF"). Se a mesma matrícula aparecer tanto em "imoveis_requerentes" quanto em
"imoveis_essenciais", repita o mesmo cartório/CRI e o mesmo proprietário nos dois lugares — é o mesmo
bem, a essencialidade não muda quem é o dono nem onde está registrado.

Responda SOMENTE com o JSON.

{
  "rj_numero": "", "vara": "",
  "requerentes": "Nome · CPF/CNPJ ... (referência)",
  "advogados_requerentes": "(referência)",
  "administrador_judicial": "Nome (referência)",
  "data_pedido": "DD/MM/AAAA (referência)", "data_deferimento": "DD/MM/AAAA (referência)",
  "consolidacao_substancial": "Deferido/Indeferido/... (referência)",
  "periodo_blindagem": "Ativo/Inativo (referência)",
  "previsao_encerramento_stay": "DD/MM/AAAA (referência)",
  "stay_prorrogavel": "Sim (referência) | Não (referência) | ''",
  "recursos_relevantes": [{"recurso": "(referência)", "status": ""}],
  "imoveis_requerentes": [{"matricula": "nº (referência)", "cartorio": "", "descricao": "", "proprietario": ""}],
  "imoveis_essenciais": [{"matricula": "nº (referência)", "cartorio": "", "descricao": "", "proprietario": ""}],
  "prj_classe_ii": {"desagio": "(referência)", "carencia": "", "parcelas": "", "juros": "", "correcao": ""},
  "prj_classe_iii": {"desagio": "(referência)", "carencia": "", "parcelas": "", "juros": "", "correcao": ""},
  "qgc": {"classe_i": "R$ (referência)", "classe_ii": "R$", "classe_iii": "R$", "classe_iv": "R$", "total": "R$"},
  "agc_situacao": "opção (referência)",
  "agc_1a": "DD/MM/AAAA (referência)", "agc_2a": "DD/MM/AAAA (referência)", "agc_continuacao": "(referência)",
  "documentos_salvos": {
    "peticao_inicial":       "referência exata onde consta, ou 'Não consta'",
    "quadro_ativos":         "referência exata onde consta, ou 'Não consta'",
    "pericia_previa":        "referência exata onde consta, ou 'Não consta'",
    "laudo_imoveis":         "referência exata onde consta, ou 'Não consta'",
    "ultimo_rma":            "referência exata onde consta, ou 'Não consta'",
    "qgc_recuperando":       "referência exata onde consta, ou 'Não consta'",
    "qgc_aj":                "referência exata onde consta, ou 'Não consta'",
    "relatorio_divergencia": "referência exata onde consta, ou 'Não consta'",
    "prj_aditivos":          "referência exata onde consta, ou 'Não consta'",
    "atas_agc":              "referência exata onde consta, ou 'Não consta'"
  }
}

Para "documentos_salvos": para cada documento, escreva a referência exata (fls./Mov./ID/Evento) de onde
ele aparece no processo, ou "Não consta" se o documento não existir nas fontes — sem colchetes de
molde de preenchimento manual, sem status de "anexado/pendente", só a referência ou "Não consta".

=== RELATÓRIO CONSOLIDADO (fonte principal) ===
"""


def _montar_fonte_rj(relatorio: str, texto_bruto: str) -> str:
    """Concatena as duas fontes do checklist de RJ: o relatório consolidado
    (fonte principal, já revisado e com referências corretas) e o texto bruto
    extraído do processo (fonte complementar, só para o que o relatório não
    cobrir)."""
    relatorio = (relatorio or "").strip()[:500_000]
    texto_bruto = (texto_bruto or "").strip()[:400_000]
    partes = [relatorio or "(relatório não disponível)"]
    if texto_bruto:
        partes.append("\n\n=== TEXTO BRUTO EXTRAÍDO (fonte complementar) ===\n" + texto_bruto)
    return "\n".join(partes)


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
    agc_situacao = str(dados.get("agc_situacao", "") or "").strip()
    if _norm_cb(agc_situacao) in ("", "nao consta"):
        agc_situacao = "Sem datas designadas"
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Situação da AGC", _cb(["Sem datas designadas", "Convocada", "1ª convocação sem quórum",
                                 "Período de suspensão", "Plano aprovado / rejeitado"],
                                agc_situacao)),
        ("1ª Convocação",              dados.get("agc_1a", "")),
        ("2ª Convocação",              dados.get("agc_2a", "")),
        ("Continuação da 2ª Convocação", dados.get("agc_continuacao", "")),
    ])
    _spacer(doc)

    # ── 6. CHECKLIST DOS DOCUMENTOS SALVOS ───────────────────────────────
    # (o modelo oficial pula do item 4 direto pro 6 — não há item 5 visível)
    _sec_title(doc, "6. CHECKLIST DOS DOCUMENTOS SALVOS")
    docs_salvos = _as_dict(dados.get("documentos_salvos"))
    rows6 = [(label, _as_str(docs_salvos.get(key, "")).strip() or "Não consta") for key, label, _opts in _DOCS_ITEM6]
    _grid_table(doc, ["DOCUMENTO", "REFERÊNCIA"], rows6, [10.9, 6.0])

    _rodape_conf(doc)
    return _salvar(doc, "Checklist_RJ", dados)


def gerar_checklist_rj(relatorio: str, texto_bruto: str, client, model: str) -> str:
    fonte = _montar_fonte_rj(relatorio, texto_bruto)
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


def _build_checklist_creditos(dados: dict, sufixo: str = "") -> str:
    doc = _base_doc()
    hoje = _date.today().strftime("%d/%m/%Y")
    _titulo(doc, "Análise de Créditos em Recuperação Judicial")
    _info_rj(doc, dados, hoje)
    _creditor_sections(doc, dados)
    _rodape_conf(doc)
    return _salvar(doc, "Analise_Creditos_RJ", dados, sufixo=sufixo)


def _extrair_cred_scoped(fonte, nome, doc, client, model) -> dict:
    foco = (
        "═══ FOCO OBRIGATÓRIO ═══\n"
        f'Extraia os dados APENAS do credor "{nome}"'
        + (f" (CPF/CNPJ: {doc})" if doc else "")
        + ". Ignore todos os outros credores do processo.\n\n"
    )
    return _extrair(foco + _PROMPT_CRED, fonte, client, model)


_PROMPT_IDENTIFICAR_EXEQUENTES = """\
Você está analisando o texto extraído de processos relacionados a um crédito (execuções, ações de
cobrança, cumprimento de sentença etc.), SEM nenhum credor-alvo especificado previamente.

Identifique o(s) CREDOR(ES) EXEQUENTE(S) — a parte que ocupa o polo ativo (exequente/credor/autor) das
execuções ou ações de cobrança encontradas no texto. NÃO confunda com o polo passivo (executado/devedor).
Se o mesmo credor aparecer em mais de uma execução, liste-o UMA ÚNICA vez.

Retorne APENAS um JSON válido (sem markdown) com esta estrutura:
{"exequentes": [{"nome": "nome/razão social completo", "cpf_cnpj": "apenas dígitos, ou null"}]}

Se não houver nenhum polo ativo identificável no texto, retorne {"exequentes": []}.

TEXTO:
"""


def _identificar_exequentes(fonte: str, client, model: str) -> list:
    """Quando nenhum credor-alvo foi informado, tenta achar o(s) exequente(s) das execuções/ações
    relacionadas — é o único sinal disponível de "quem é o credor" no fluxo sem RJ (só relacionados)."""
    resultado = _extrair(_PROMPT_IDENTIFICAR_EXEQUENTES, fonte, client, model)
    exequentes = resultado.get("exequentes") if isinstance(resultado, dict) else None
    if not isinstance(exequentes, list):
        return []
    vistos, candidatos = set(), []
    for ex in exequentes:
        if not isinstance(ex, dict):
            continue
        nome = str(ex.get("nome") or "").strip()
        if not nome:
            continue
        chave = re.sub(r"\s+", " ", nome).strip().lower()
        if chave in vistos:
            continue
        vistos.add(chave)
        candidatos.append((nome, str(ex.get("cpf_cnpj") or "").strip()))
    return candidatos


def _gerar_docs_por_credor(fonte: str, creditores: list, client, model: str) -> list:
    caminhos = []
    for nome, doc_id in creditores:
        dados = _extrair_cred_scoped(fonte, nome, doc_id, client, model)
        if not (dados.get("credor") or "").strip():
            dados["credor"] = nome + (f" · {doc_id}" if doc_id else "")
        caminhos.append(_build_checklist_creditos(dados, sufixo=nome))
    return caminhos


def gerar_checklist_creditos(fonte: str, client, model: str, creditores=None):
    # Credores informados manualmente: um .docx por credor, escopado à extração de cada um
    if creditores:
        return _gerar_docs_por_credor(fonte, creditores, client, model)
    # Sem credor informado: tenta identificar o(s) credor(es) exequente(s) das execuções/ações
    # relacionadas antes de cair no modo genérico — essencial no fluxo sem RJ (só processos
    # relacionados), em que o único sinal de "quem é o credor" é o polo ativo das execuções.
    candidatos = _identificar_exequentes(fonte, client, model)
    if candidatos:
        return _gerar_docs_por_credor(fonte, candidatos, client, model)
    # Fallback: não foi possível identificar nenhum exequente — extração genérica de sempre
    dados = _extrair(_PROMPT_CRED, fonte, client, model)
    return _build_checklist_creditos(dados)
