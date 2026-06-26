# -*- coding: utf-8 -*-
"""Geração do Dossiê PPA Invista em Word, no formato Parecer_Invista_PPA_v2."""

import json
import os
import re
import tempfile
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from google.genai import types

from utils import _retry

# ── Paleta Invista (extraída do template) ─────────────────────────────────
_LARANJA   = "E8440A"   # fundo de células label
_BRANCO    = "FFFFFF"
_CREME     = "FAFAF9"   # fundo alternado
_CREME_LT  = "FEF0EB"   # destaque suave

# ── Logo Invista ──────────────────────────────────────────────────────────
_LOGO_PATH = Path(__file__).parent / "assets" / "invista_logo.png"


# ── helpers python-docx ───────────────────────────────────────────────────

def _set_bg(cell, hex6: str):
    tcp = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex6)
    tcp.append(shd)


def _set_borders(cell, color="DDDDDD", sz=4):
    tcp = cell._tc.get_or_add_tcPr()
    be = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
        be.append(b)
    tcp.append(be)


def _cell_pad(cell, top=60, left=100, bottom=60, right=100):
    tcp = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tcp.append(mar)


def _write(cell, text: str, bold=False, pt=9, white=False, italic=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if align:
        p.alignment = align
    r = p.add_run(str(text or ""))
    r.font.name = "Calibri"
    r.font.size = Pt(pt)
    r.bold = bold
    r.italic = italic
    if white:
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _label(cell, text: str):
    _set_bg(cell, _LARANJA)
    _set_borders(cell)
    _cell_pad(cell)
    _write(cell, text, bold=True, white=True)


def _value(cell, text: str, cream=False):
    _set_bg(cell, _CREME if cream else _BRANCO)
    _set_borders(cell)
    _cell_pad(cell)
    _write(cell, text)


def _full_span_row(table, text: str):
    """Linha laranja que abrange todas as colunas — usada como separador de bloco."""
    row = table.add_row()
    n = len(row.cells)
    cell = row.cells[0]
    for i in range(1, n):
        cell = cell.merge(row.cells[i])
    _set_bg(cell, _LARANJA)
    _set_borders(cell, color="E8440A")
    _cell_pad(cell)
    _write(cell, text, bold=True, white=True)
    return row


def _two_col_table(doc, rows_data: list, header: str | None = None,
                   w_label=5.0, w_value=11.5):
    """
    Cria tabela 2 colunas: label | value.
    rows_data = [(label, value), ...]
    """
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    if header:
        _full_span_row(t, header)
    for label, value in rows_data:
        row = t.add_row()
        _label(row.cells[0])
        # set label text
        _set_bg(row.cells[0], _LARANJA)
        _set_borders(row.cells[0])
        _cell_pad(row.cells[0])
        _write(row.cells[0], label, bold=True, white=True)
        _value(row.cells[1], value)
    # widths
    for row in t.rows:
        row.cells[0].width = Cm(w_label)
        row.cells[1].width = Cm(w_value)
    return t


def _section_title(doc, text: str, spacing_before=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(spacing_before)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Calibri"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0xE8, 0x44, 0x0A)


def _sub_title(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Calibri"
    r.font.size = Pt(10)


def _body_text(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)


def _spacer(doc, n=1):
    for _ in range(n):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)


# ── Prompt de extração de dados para o dossiê ─────────────────────────────

_PROMPT_DOSSIE = """\
Com base no relatório jurídico abaixo, extraia os dados para preencher um dossiê no formato JSON.
Extraia APENAS o que está explicitamente mencionado no relatório. Use string vazia "" para campos ausentes.
Responda SOMENTE com o JSON, sem texto adicional.

Formato esperado:
{
  "nome_caso": "nome identificador curto do caso (ex: BASF x São Lourenço)",
  "data_analise": "DD/MM/AAAA",
  "exequentes": "Nome · CNPJ: 00.000.000/0001-00",
  "executados": "Nome1 · CPF/CNPJ: ...; Nome2 · CPF/CNPJ: ...",
  "sat_total": "R$ 0.000.000,00",
  "teses_principais": "Penhora Direta / IDPJ / Fraude à Execução",
  "risco_juridico": "resumo dos principais riscos identificados",
  "consideracoes_gerais": "2-4 parágrafos de análise geral do caso",
  "creditos": [
    {
      "id": "Crédito [Credor]",
      "numero_processo": "0000000-00.0000.0.00.0000",
      "vara_comarca": "ex: 1ª Vara Cível — Cascavel/PR",
      "exequente_info": "Nome · Adv: Escritório | OAB/XX 00000",
      "executados_info": "Nome · Adv: Escritório | OAB/XX 00000",
      "data_distribuicao": "DD/MM/AAAA",
      "sop": "R$ 0,00",
      "sat": "R$ 0,00",
      "honorarios": "00%",
      "lastro": "CCB nº / Contrato nº",
      "data_emissao": "DD/MM/AAAA",
      "data_vencimento": "DD/MM/AAAA",
      "assinaturas": "Nome A, Nome B",
      "garantia": "descrição da garantia",
      "status_processo": "ex: Em fase de penhora de imóvel",
      "ind_cm": "ex: IPCA",
      "ind_jr": "ex: 12% a.a.",
      "ind_jm": "ex: 1% a.m.",
      "ind_multa": "ex: 2%",
      "ind_cap": "ex: Mensal",
      "plan_cm": "",
      "plan_jr": "",
      "plan_multa": "",
      "plan_cap": "",
      "plan_obs": "",
      "citacoes": [
        {"executado": "Nome", "modalidade": "AR/OJ/Carta Precatória", "data": "DD/MM/AAAA", "fls": "00"}
      ],
      "embargos": [
        {
          "numero": "nº do processo",
          "embargante": "Nome",
          "data_dist": "DD/MM/AAAA",
          "tese": "resumo da tese dos embargos",
          "andamentos_resumo": "principais andamentos",
          "status": "ex: Julgado improcedente"
        }
      ],
      "recursos": [
        {
          "numero": "nº ou ID do recurso",
          "recorrente": "Nome",
          "decisao_recorrida": "ex: Sentença de improcedência dos embargos",
          "data_dist": "DD/MM/AAAA",
          "tese": "resumo da tese",
          "andamentos_resumo": "principais andamentos",
          "status": "ex: Pendente de julgamento"
        }
      ],
      "constricoes": [
        {"tipo": "Penhora de imóvel/Sisbajud/outro", "descricao": "detalhes", "valor": "R$ 0,00", "status": "Ativa"}
      ],
      "andamentos": [
        {"data": "DD/MM/AAAA", "descricao": "descrição objetiva", "fls": "Mov. X | fls. XX"}
      ]
    }
  ]
}

RELATÓRIO:
"""


# ── Extração JSON via Gemini ──────────────────────────────────────────────

def _extrair_dados(relatorio: str, client, model: str) -> dict:
    prompt = _PROMPT_DOSSIE + relatorio[:60_000]
    try:
        config = types.GenerateContentConfig(response_mime_type="application/json")
        def _fn():
            return client.models.generate_content(
                model=model,
                contents=[types.Content(role="user", parts=[types.Part(text=prompt)])],
                config=config,
            ).text
        raw = _retry(_fn, tentativas=3, espera_base=10)
        # Limpar eventual markdown fencing
        raw = re.sub(r"^```[a-z]*\n?", "", raw.strip(), flags=re.IGNORECASE)
        raw = re.sub(r"\n?```$", "", raw.strip())
        return json.loads(raw)
    except Exception:
        return {}


# ── Construção do documento Word ──────────────────────────────────────────

def _build_doc(dados: dict) -> str:
    doc = Document()

    # Página A4, margens 2 cm
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, attr, Cm(2))

    # Fonte padrão
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    # ── Cabeçalho — logo + título ─────────────────────────────────────────
    try:
        import io
        logo_stream = io.BytesIO(_LOGO_PATH.read_bytes()) if _LOGO_PATH.exists() else None
        if logo_stream is None:
            raise FileNotFoundError("logo not found")
        t_hdr = doc.add_table(rows=1, cols=2)
        t_hdr.style = "Table Grid"
        # remover bordas da tabela do cabeçalho
        for cell in t_hdr.rows[0].cells:
            tcp = cell._tc.get_or_add_tcPr()
            be = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), "none")
                be.append(b)
            tcp.append(be)
        # Logo (col 0)
        c0 = t_hdr.rows[0].cells[0]
        c0.width = Cm(4)
        p_logo = c0.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r_logo = p_logo.add_run()
        r_logo.add_picture(logo_stream, width=Cm(3.5))
        # Título (col 1)
        c1 = t_hdr.rows[0].cells[1]
        c1.width = Cm(12.5)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r1 = p1.add_run("Dossiê | Recuperação de Crédito")
        r1.bold = True
        r1.font.name = "Calibri"
        r1.font.size = Pt(14)
        r1.font.color.rgb = RGBColor(0xE8, 0x44, 0x0A)
        p2 = c1.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r2 = p2.add_run("Análise de Crédito · PPA Invista")
        r2.font.name = "Calibri"
        r2.font.size = Pt(9)
        r2.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
    except Exception:
        p = doc.add_paragraph("Dossiê | Recuperação de Crédito  ·  Análise de Crédito · PPA Invista")
        p.runs[0].bold = True
        p.runs[0].font.color.rgb = RGBColor(0xE8, 0x44, 0x0A)

    # Linha divisória laranja
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(4)
    p_div.paragraph_format.border_bottom = None
    pfmt = p_div._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "12")
    bot.set(qn("w:space"), "1")
    bot.set(qn("w:color"), "E8440A")
    pb.append(bot)
    pfmt.append(pb)

    # ── Info box ─────────────────────────────────────────────────────────
    from datetime import date as _date
    hoje = _date.today().strftime("%d/%m/%Y")
    _two_col_table(doc, [
        ("Nome do Caso",               dados.get("nome_caso", "")),
        ("Data da Análise",             dados.get("data_analise", hoje)),
        ("Advogada(o) Responsável",     ""),
    ], w_label=5.5, w_value=11.0)

    _spacer(doc)

    # ── Seção 0 — Visão Geral ─────────────────────────────────────────────
    _section_title(doc, "0. VISÃO GERAL DO CASO")

    _two_col_table(doc, [
        ("Exequente(s)",                dados.get("exequentes", "")),
        ("Executado(s)",                dados.get("executados", "")),
        ("SAT",                         dados.get("sat_total", "R$")),
        ("Total atingível mapeado — VM","R$"),
        ("Total atingível mapeado — VP","R$"),
        ("Passivo total identificado",  "R$"),
        ("Tese(s) principal(is)",       dados.get("teses_principais", "")),
        ("Risco Jurídico",              dados.get("risco_juridico", "")),
    ], w_label=5.5, w_value=11.0)

    _spacer(doc)

    # Quadro ativos consolidados
    _sub_title(doc, "VISÃO CONSOLIDADA DOS ATIVOS")
    t_ativos = doc.add_table(rows=0, cols=4)
    t_ativos.style = "Table Grid"
    # Header
    hdr = t_ativos.add_row()
    for cell, txt in zip(hdr.cells, ["TESE", "VM (R$)", "VP (R$)", "OBSERVAÇÕES"]):
        _label(cell)
        _write(cell, txt, bold=True, white=True)
    # Linhas
    for tese in ["Penhora Direta", "IDPJ", "Fraude à Execução", "Outras Teses", "TOTAL GERAL"]:
        row = t_ativos.add_row()
        _value(row.cells[0], tese, cream=(tese != "TOTAL GERAL"))
        _value(row.cells[1], "R$", cream=(tese != "TOTAL GERAL"))
        _value(row.cells[2], "R$", cream=(tese != "TOTAL GERAL"))
        _value(row.cells[3], "", cream=(tese != "TOTAL GERAL"))
        if tese == "TOTAL GERAL":
            for c in row.cells:
                _set_bg(c, _CREME_LT)
    # Larguras
    for row in t_ativos.rows:
        row.cells[0].width = Cm(4.5)
        row.cells[1].width = Cm(3.0)
        row.cells[2].width = Cm(3.0)
        row.cells[3].width = Cm(6.0)

    _spacer(doc)

    # Principais considerações
    _sub_title(doc, "PRINCIPAIS CONSIDERAÇÕES")
    cons = dados.get("consideracoes_gerais", "")
    if cons:
        for para in cons.split("\n"):
            if para.strip():
                _body_text(doc, para.strip())
    else:
        _body_text(doc, "")

    # ── Seção 1 — Visão Jurídica ──────────────────────────────────────────
    _section_title(doc, "1. VISÃO JURÍDICA")

    creditos = dados.get("creditos") or []
    if not creditos:
        _body_text(doc, "Nenhum crédito identificado.")
    else:
        for idx, cred in enumerate(creditos, 1):
            cid = cred.get("id") or f"Crédito {idx}"
            _sub_title(doc, f"1.{idx} {cid}")

            # Dados do Processo
            _two_col_table(doc, [
                ("Número do processo",   cred.get("numero_processo", "")),
                ("Vara / Comarca",       cred.get("vara_comarca", "")),
                ("Exequente",            cred.get("exequente_info", "")),
                ("Executado(s)",         cred.get("executados_info", "")),
                ("Data de distribuição", cred.get("data_distribuicao", "")),
                ("SOP",                  cred.get("sop", "R$")),
                ("SAT",                  cred.get("sat", "R$")),
                ("Honorários",           cred.get("honorarios", "")),
                ("Lastro / Instrumento", cred.get("lastro", "")),
                ("Data de Emissão",      cred.get("data_emissao", "")),
                ("Data do Vencimento",   cred.get("data_vencimento", "")),
                ("Assinaturas",          cred.get("assinaturas", "")),
                ("Garantia",             cred.get("garantia", "")),
                ("Status do Processo",   cred.get("status_processo", "")),
            ], header="Dados do Processo", w_label=5.5, w_value=11.0)

            _spacer(doc)

            # Índices de Correção do Contrato
            _two_col_table(doc, [
                ("Correção monetária",    cred.get("ind_cm", "")),
                ("Juros remuneratórios",  cred.get("ind_jr", "")),
                ("Juros moratórios",      cred.get("ind_jm", "")),
                ("Multa moratória",       cred.get("ind_multa", "")),
                ("Capitalização",         cred.get("ind_cap", "")),
            ], header="Índices de Correção do Contrato", w_label=5.5, w_value=11.0)

            _spacer(doc)

            # Planilha Inicial
            _two_col_table(doc, [
                ("Correção monetária",    cred.get("plan_cm", "")),
                ("Juros remuneratórios",  cred.get("plan_jr", "")),
                ("Multa moratória",       cred.get("plan_multa", "")),
                ("Capitalização",         cred.get("plan_cap", "")),
                ("Observações",           cred.get("plan_obs", "")),
            ], header="Planilha Inicial", w_label=5.5, w_value=11.0)

            _spacer(doc)

            # Citação
            citacoes = cred.get("citacoes") or []
            t_cit = doc.add_table(rows=0, cols=4)
            t_cit.style = "Table Grid"
            _full_span_row(t_cit, "Citação")
            hdr_cit = t_cit.add_row()
            for cell, txt in zip(hdr_cit.cells, ["EXECUTADO", "MODALIDADE", "DATA", "FLS."]):
                _label(cell)
                _write(cell, txt, bold=True, white=True)
            if citacoes:
                for cit in citacoes:
                    r = t_cit.add_row()
                    _value(r.cells[0], cit.get("executado", ""))
                    _value(r.cells[1], cit.get("modalidade", ""))
                    _value(r.cells[2], cit.get("data", ""))
                    _value(r.cells[3], cit.get("fls", ""))
            else:
                r = t_cit.add_row()
                for c in r.cells:
                    _value(c, "")
            for row in t_cit.rows:
                row.cells[0].width = Cm(5.0)
                row.cells[1].width = Cm(4.5)
                row.cells[2].width = Cm(3.0)
                row.cells[3].width = Cm(4.0)

            _spacer(doc)

            # Embargos
            embargos = cred.get("embargos") or []
            if embargos:
                for ei, emb in enumerate(embargos, 1):
                    _two_col_table(doc, [
                        ("Nº do Processo",        emb.get("numero", "")),
                        ("Embargante",            emb.get("embargante", "")),
                        ("Data da Distribuição",  emb.get("data_dist", "")),
                        ("Tese dos Embargos",     emb.get("tese", "")),
                        ("Principais Andamentos", emb.get("andamentos_resumo", "")),
                        ("Status Atual",          emb.get("status", "")),
                    ], header=f"Embargo nº {ei}", w_label=5.5, w_value=11.0)
                    _spacer(doc)

            # Recursos
            recursos = cred.get("recursos") or []
            if recursos:
                for ri, rec in enumerate(recursos, 1):
                    _two_col_table(doc, [
                        ("Nº do Processo",        rec.get("numero", "")),
                        ("Recorrente",            rec.get("recorrente", "")),
                        ("Decisão Recorrida",     rec.get("decisao_recorrida", "")),
                        ("Data da Distribuição",  rec.get("data_dist", "")),
                        ("Tese do Recurso",       rec.get("tese", "")),
                        ("Principais Andamentos", rec.get("andamentos_resumo", "")),
                        ("Status Atual",          rec.get("status", "")),
                    ], header=f"Recurso nº {ri}", w_label=5.5, w_value=11.0)
                    _spacer(doc)

            # Constrições Vigentes
            constricoes = cred.get("constricoes") or []
            t_con = doc.add_table(rows=0, cols=4)
            t_con.style = "Table Grid"
            _full_span_row(t_con, "Constrições Vigentes")
            hdr_con = t_con.add_row()
            for cell, txt in zip(hdr_con.cells, ["TIPO", "DESCRIÇÃO", "VALOR (R$)", "STATUS"]):
                _label(cell)
                _write(cell, txt, bold=True, white=True)
            if constricoes:
                for con in constricoes:
                    r = t_con.add_row()
                    _value(r.cells[0], con.get("tipo", ""))
                    _value(r.cells[1], con.get("descricao", ""))
                    _value(r.cells[2], con.get("valor", ""))
                    _value(r.cells[3], con.get("status", ""))
            else:
                r = t_con.add_row()
                for c in r.cells:
                    _value(c, "")
            for row in t_con.rows:
                row.cells[0].width = Cm(3.5)
                row.cells[1].width = Cm(5.5)
                row.cells[2].width = Cm(3.5)
                row.cells[3].width = Cm(4.0)

            _spacer(doc)

            # Principais Andamentos
            andamentos = cred.get("andamentos") or []
            t_and = doc.add_table(rows=0, cols=3)
            t_and.style = "Table Grid"
            _full_span_row(t_and, "Principais Andamentos Processuais")
            hdr_and = t_and.add_row()
            for cell, txt in zip(hdr_and.cells, ["DATA", "DESCRIÇÃO", "FLS. / EVENTO"]):
                _label(cell)
                _write(cell, txt, bold=True, white=True)
            if andamentos:
                for an in andamentos:
                    r = t_and.add_row()
                    _value(r.cells[0], an.get("data", ""))
                    _value(r.cells[1], an.get("descricao", ""))
                    _value(r.cells[2], an.get("fls", ""))
            else:
                r = t_and.add_row()
                for c in r.cells:
                    _value(c, "")
            for row in t_and.rows:
                row.cells[0].width = Cm(2.5)
                row.cells[1].width = Cm(11.5)
                row.cells[2].width = Cm(2.5)

            _spacer(doc)

    # ── Seção 2 — Teses de Recuperação (placeholder) ──────────────────────
    _section_title(doc, "2. TESES DE RECUPERAÇÃO")
    for sub in [
        "2.1 Penhora Direta",
        "2.2 Incidente de Desconsideração da Personalidade Jurídica",
        "2.3 Fraude à Execução",
        "2.4 Outras teses",
    ]:
        _sub_title(doc, sub)
        _body_text(doc, "")
        _spacer(doc)

    # ── Seção 3 — Passivo (placeholder) ───────────────────────────────────
    _section_title(doc, "3. PASSIVO")

    for sub, cols in [
        ("3.1 Passivo Fiscal — e-CAC",      ["NOME", "CPF / CNPJ", "SALDO e-CAC (R$)"]),
        ("3.2 Passivo Trabalhista",          ["Nº CNJ", "Vinculado a", "Distribuição", "Autor", "Valor Causa (R$)", "SAT Est. (R$)", "Status"]),
        ("3.3 Passivo Cível (Terceiros)",    ["Nº CNJ", "Vinculado a", "Distribuição", "Exequente", "Valor Causa (R$)", "SAT Est. (R$)", "Obs."]),
    ]:
        _sub_title(doc, sub)
        t = doc.add_table(rows=2, cols=len(cols))
        t.style = "Table Grid"
        for c, txt in zip(t.rows[0].cells, cols):
            _label(c)
            _write(c, txt, bold=True, white=True)
        for c in t.rows[1].cells:
            _value(c, "")
        _spacer(doc)

    # ── Seção 5 — Pendências ──────────────────────────────────────────────
    _section_title(doc, "5. PENDÊNCIAS")
    t_pend = doc.add_table(rows=2, cols=3)
    t_pend.style = "Table Grid"
    for c, txt in zip(t_pend.rows[0].cells, ["PENDÊNCIA", "PROTOCOLO", "PRAZO / STATUS"]):
        _label(c)
        _write(c, txt, bold=True, white=True)
    for c in t_pend.rows[1].cells:
        _value(c, "")
    for row in t_pend.rows:
        row.cells[0].width = Cm(8.5)
        row.cells[1].width = Cm(4.0)
        row.cells[2].width = Cm(4.0)

    _spacer(doc)

    # ── Seção 6 — Elaboração e Revisão ────────────────────────────────────
    _section_title(doc, "6. ELABORAÇÃO E REVISÃO")
    t_rev = doc.add_table(rows=0, cols=3)
    t_rev.style = "Table Grid"
    hdr_rev = t_rev.add_row()
    for c, txt in zip(hdr_rev.cells, ["FUNÇÃO", "NOME", "DATA"]):
        _label(c)
        _write(c, txt, bold=True, white=True)
    for func in ["Advogada(o) Responsável pela Análise", "Advogada(o) Revisor"]:
        r = t_rev.add_row()
        _value(r.cells[0], func, cream=True)
        _value(r.cells[1], "")
        _value(r.cells[2], "")
    for row in t_rev.rows:
        row.cells[0].width = Cm(8.0)
        row.cells[1].width = Cm(5.5)
        row.cells[2].width = Cm(3.0)

    _spacer(doc)

    # Rodapé confidencialidade (parágrafo no corpo)
    p_rod = doc.add_paragraph()
    p_rod.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_rod = p_rod.add_run(
        "Documento confidencial. Uso interno exclusivo do Time PPA — Invista. "
        "Vedada a reprodução ou distribuição a terceiros sem autorização prévia."
    )
    r_rod.italic = True
    r_rod.font.name = "Calibri"
    r_rod.font.size = Pt(8)
    r_rod.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    # Salvar
    nome_caso = re.sub(r"[^\w\s-]", "", dados.get("nome_caso", "caso")).strip().replace(" ", "_")[:40]
    caminho = os.path.join(tempfile.gettempdir(), f"dossie_ppa_{nome_caso}.docx")
    doc.save(caminho)
    return caminho


# ── Entry point ───────────────────────────────────────────────────────────

def gerar_dossie_word(relatorio: str, client, model: str) -> str:
    """Extrai dados do relatório via Gemini e gera o Word no formato PPA."""
    dados = _extrair_dados(relatorio, client, model)
    return _build_doc(dados)
