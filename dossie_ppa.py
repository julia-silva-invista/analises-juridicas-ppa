# -*- coding: utf-8 -*-
"""Geração do Dossiê PPA Invista em Word — modelo Parecer_Invista_PPA_v2_Atualizada 3.0."""

import io
import json
import os
import re
import tempfile
from datetime import date as _date
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

from google.genai import types

from utils import _retry

# ── Paleta Invista (extraída do template v3.0) ────────────────────────────
_LARANJA   = "E8440A"   # headers de tabela, títulos
_CINZA     = "F5F5F3"   # coluna-label (embargo/recurso), linhas TOTAL
_BRANCO    = "FFFFFF"   # células de valor
_DESTAQUE  = "FEF0EB"   # caixa de considerações
_BORDA     = "E8E8E8"   # bordas suaves
_TXT       = "555555"   # texto escuro padrão
_TXT_MUTE  = "AAAAAA"   # notas/placeholder

_LOGO_PATH = Path(__file__).parent / "assets" / "invista_logo.png"


# ══════════════════════════════════════════════════════════════════════════
# Helpers de célula (baixo nível)
# ══════════════════════════════════════════════════════════════════════════

def _set_bg(cell, hex6: str):
    tcp = cell._tc.get_or_add_tcPr()
    for old in tcp.findall(qn("w:shd")):
        tcp.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex6)
    tcp.append(shd)


def _set_borders(cell, color=_BORDA, sz=4):
    tcp = cell._tc.get_or_add_tcPr()
    for old in tcp.findall(qn("w:tcBorders")):
        tcp.remove(old)
    be = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(sz))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), color)
        be.append(b)
    tcp.append(be)


def _cell_pad(cell, top=50, left=90, bottom=50, right=90):
    tcp = cell._tc.get_or_add_tcPr()
    for old in tcp.findall(qn("w:tcMar")):
        tcp.remove(old)
    mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(val))
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tcp.append(mar)


def _apply_font(run, bold, size, color, italic):
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    try:
        run.font.color.rgb = RGBColor.from_string(color)
    except Exception:
        pass


def _write(cell, text, bold=False, size=9, color=_TXT, italic=False, align=None, valign=True):
    cell.text = ""
    if valign:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    if align:
        p.alignment = align
    parts = str(text if text is not None else "").split("\n")
    for j, line in enumerate(parts):
        r = p.add_run(line)
        _apply_font(r, bold, size, color, italic)
        if j < len(parts) - 1:
            r.add_break()


# ══════════════════════════════════════════════════════════════════════════
# Helpers de linha / tabela (médio nível)
# ══════════════════════════════════════════════════════════════════════════

def _widths(row, ws):
    for i, w in enumerate(ws):
        row.cells[i].width = Cm(w)


def _orange_header(table, labels, ws=None):
    """Linha de header laranja com texto branco em negrito."""
    row = table.add_row()
    for i, lab in enumerate(labels):
        c = row.cells[i]
        _set_bg(c, _LARANJA); _set_borders(c); _cell_pad(c)
        _write(c, lab, bold=True, size=9, color=_BRANCO)
    if ws:
        _widths(row, ws)
    return row


def _span_header(table, text):
    """Header laranja que abrange todas as colunas (ex.: DADOS DA MEMÓRIA DE CÁLCULO)."""
    row = table.add_row()
    n = len(row.cells)
    cell = row.cells[0]
    for i in range(1, n):
        cell = cell.merge(row.cells[i])
    _set_bg(cell, _LARANJA); _set_borders(cell); _cell_pad(cell)
    _write(cell, text, bold=True, size=9, color=_BRANCO)
    return row


def _kv_table(doc, header, rows, w_label=5.9, w_value=11.0):
    """Tabela 2 col estilo 'header laranja' (CAMPO|INFORMAÇÃO) + labels/valores brancos."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if header:
        if isinstance(header, str):
            _span_header(t, header)
        else:
            _orange_header(t, header, [w_label, w_value])
    for label, value in rows:
        r = t.add_row()
        cl, cv = r.cells[0], r.cells[1]
        _set_bg(cl, _BRANCO); _set_borders(cl); _cell_pad(cl)
        _write(cl, label, bold=True, color=_TXT)
        _set_bg(cv, _BRANCO); _set_borders(cv); _cell_pad(cv)
        _write(cv, value, color=_TXT)
        _widths(r, [w_label, w_value])
    return t


def _kv_label_table(doc, rows, w_label=5.9, w_value=11.0):
    """Tabela 2 col com coluna-label cinza (embargo/recurso), sem header laranja."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for label, value in rows:
        r = t.add_row()
        cl, cv = r.cells[0], r.cells[1]
        _set_bg(cl, _CINZA); _set_borders(cl); _cell_pad(cl)
        _write(cl, label, bold=True, color=_TXT)
        _set_bg(cv, _BRANCO); _set_borders(cv); _cell_pad(cv)
        _write(cv, value, color=_TXT)
        _widths(r, [w_label, w_value])
    return t


def _grid_table(doc, headers, rows, ws, total_row=None, min_rows=1):
    """Tabela multi-coluna: header laranja + linhas brancas (+ TOTAL cinza opcional)."""
    t = doc.add_table(rows=0, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _orange_header(t, headers, ws)
    data = list(rows)
    while len(data) < min_rows:
        data.append([""] * len(headers))
    for row_data in data:
        r = t.add_row()
        for i in range(len(headers)):
            c = r.cells[i]
            val = row_data[i] if i < len(row_data) else ""
            _set_bg(c, _BRANCO); _set_borders(c); _cell_pad(c)
            _write(c, val)
        _widths(r, ws)
    if total_row:
        r = t.add_row()
        for i in range(len(headers)):
            c = r.cells[i]
            val = total_row[i] if i < len(total_row) else ""
            _set_bg(c, _CINZA); _set_borders(c); _cell_pad(c)
            _write(c, val, bold=True)
        _widths(r, ws)
    return t


# ══════════════════════════════════════════════════════════════════════════
# Helpers de parágrafo (títulos, notas, texto)
# ══════════════════════════════════════════════════════════════════════════

def _para(doc, text, bold=False, size=9.5, color=_TXT, italic=False,
          before=0, after=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if align:
        p.alignment = align
    r = p.add_run(str(text or ""))
    _apply_font(r, bold, size, color, italic)
    return p


def _sec_title(doc, text):
    return _para(doc, text, bold=True, size=13, color=_LARANJA, before=14, after=5)


def _sub_orange(doc, text):
    return _para(doc, text, bold=True, size=11, color=_LARANJA, before=10, after=4)


def _sub_gray(doc, text):
    return _para(doc, text, bold=True, size=10, color=_TXT, before=7, after=3)


def _note(doc, text):
    return _para(doc, text, size=8, color=_TXT_MUTE, italic=True, before=2, after=6)


def _body(doc, text):
    return _para(doc, text, size=9.5, color=_TXT, before=0, after=5)


def _guidance(doc, text):
    """Caixa de orientação com fundo FEF0EB (célula única)."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = t.rows[0].cells[0]
    _set_bg(c, _DESTAQUE); _set_borders(c, color="F3D9CF"); _cell_pad(c, top=90, bottom=90)
    _write(c, text, size=9, color=_TXT, italic=True)
    c.width = Cm(16.9)
    return t


def _spacer(doc, pts=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pts)


# ══════════════════════════════════════════════════════════════════════════
# Extração de dados via Gemini
# ══════════════════════════════════════════════════════════════════════════

_PROMPT_DOSSIE = """\
Com base no relatório jurídico abaixo, extraia os dados para preencher um dossiê no formato JSON.
Extraia APENAS o que está explicitamente no relatório. Use string vazia "" para o que não constar.
NÃO invente números, nomes ou datas. Responda SOMENTE com o JSON, sem texto adicional.

{
  "nome_caso": "identificador curto (ex: BASF x São Lourenço)",
  "data_analise": "DD/MM/AAAA ou vazio",
  "exequentes": "Nome · CNPJ ...",
  "executados": "Nome1 · CPF/CNPJ ...; Nome2 ...",
  "sat_total": "R$ ...",
  "passivo_fiscal": "R$ ... ou vazio",
  "passivo_trabalhista": "R$ ... ou vazio",
  "passivo_civel": "R$ ... ou vazio",
  "teses_principais": "Penhora Direta / IDPJ / Fraude à Execução",
  "risco_juridico": "resumo dos principais riscos",
  "consideracoes_gerais": "2-4 parágrafos de análise geral (separe parágrafos com \\n)",
  "creditos": [
    {
      "id": "Crédito [Credor]",
      "numero_processo": "0000000-00.0000.0.00.0000",
      "vara_comarca": "1ª Vara Cível — Cidade/UF",
      "exequente_info": "Nome · Adv: Escritório | OAB",
      "executados_info": "Nome · Adv: Escritório | OAB",
      "data_distribuicao": "DD/MM/AAAA",
      "sop": "R$ ...", "sat": "R$ ...", "criterio_sat": "ex: INPC + 12% a.a JM",
      "honorarios": "ex: 10%",
      "lastro": "CCB nº / Contrato nº", "data_emissao": "DD/MM/AAAA", "data_vencimento": "DD/MM/AAAA",
      "assinaturas": "Nomes / Fls.", "garantia": "descrição da garantia",
      "status_processo": "ex: em fase de penhora",
      "ind_cm": "", "ind_jr": "", "ind_jm": "", "ind_multa": "", "ind_cap": "",
      "plan_cm": "", "plan_jr": "", "plan_multa": "", "plan_cap": "", "plan_ponderacoes": "",
      "memoria_data_juntada": "", "memoria_total": "", "memoria_data_base": "",
      "memoria_indices": "", "memoria_ponderacoes": "",
      "citacoes": [{"executado": "", "modalidade": "AR/OJ/Precatória", "data": "", "fls": ""}],
      "embargos": [{"embargante": "", "data_dist": "", "tese": "", "andamentos_resumo": "", "status": ""}],
      "recursos": [{"recorrente": "", "decisao_recorrida": "", "data_dist": "", "tese": "", "andamentos_resumo": "", "status": ""}],
      "constricoes": [{"tipo": "", "descricao": "", "valor": "R$ ...", "status": "Ativa"}],
      "andamentos": [{"data": "", "descricao": "", "fls": ""}]
    }
  ]
}

RELATÓRIO:
"""


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
        raw = re.sub(r"^```[a-z]*\n?", "", raw.strip(), flags=re.IGNORECASE)
        raw = re.sub(r"\n?```$", "", raw.strip())
        return json.loads(raw)
    except Exception:
        return {}


# ══════════════════════════════════════════════════════════════════════════
# Construção do documento
# ══════════════════════════════════════════════════════════════════════════

def _montar_cabecalho_rodape(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(1.4)
    sec.bottom_margin = Cm(1.4)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.8)

    # Logo no cabeçalho de página (repete)
    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.space_after = Pt(0)
    if _LOGO_PATH.exists():
        try:
            hp.add_run().add_picture(str(_LOGO_PATH), width=Cm(3.2))
        except Exception:
            pass

    # Rodapé: confidencialidade + página
    footer = sec.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("Confidencial  ·  Uso interno  ·  Time PPA  ·  Invista")
    _apply_font(r, bold=False, size=8, color=_TXT_MUTE, italic=False)


def _build_doc(dados: dict) -> str:
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"
    doc.styles["Normal"].font.size = Pt(9.5)

    _montar_cabecalho_rodape(doc)

    hoje = _date.today().strftime("%d/%m/%Y")

    # ── Título ────────────────────────────────────────────────────────────
    _para(doc, "ANÁLISE DE CRÉDITO", bold=True, size=21, color=_LARANJA, before=2, after=0)
    _para(doc, "PPA Invista", bold=False, size=11, color=_TXT_MUTE, before=0, after=8)

    # ── Info box ─────────────────────────────────────────────────────────
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
        ("Nome do Caso",           dados.get("nome_caso", "")),
        ("Data da Análise",        dados.get("data_analise") or hoje),
        ("Advogada Responsável",   ""),
    ])
    _spacer(doc)

    # ══ 0. VISÃO GERAL DO CASO ═══════════════════════════════════════════
    _sec_title(doc, "0. VISÃO GERAL DO CASO")

    pf = dados.get("passivo_fiscal", "")
    pt = dados.get("passivo_trabalhista", "")
    pc = dados.get("passivo_civel", "")
    passivo_cell = (
        f"Fiscal: {pf or 'R$ [___]'}\n"
        f"Trabalhista: {pt or 'R$ [___]'}\n"
        f"Cível: {pc or 'R$ [___]'}\n"
        f"TOTAL: R$ [___]"
    )
    _kv_table(doc, ["INDICADOR", "VALOR / INFORMAÇÃO"], [
        ("Exequente(s)",                 dados.get("exequentes", "")),
        ("Executado(s)",                 dados.get("executados", "")),
        ("SAT",                          dados.get("sat_total", "R$")),
        ("Total atingível mapeado — VM", "R$ [___]"),
        ("Total atingível mapeado — VP", "R$ [___]"),
        ("Passivo total identificado",   passivo_cell),
        ("Tese(s) principal(is)",        dados.get("teses_principais", "")),
        ("Risco Jurídico",               dados.get("risco_juridico", "")),
    ])
    _spacer(doc)

    # Visão consolidada dos ativos
    _sub_orange(doc, "VISÃO CONSOLIDADA DOS ATIVOS")
    _note(doc, "Quadro-resumo dos ativos atingíveis por tese. Os valores já descontam os ônus incidentes sobre os ativos.")
    _grid_table(
        doc,
        ["TESE", "VM (R$)", "VP (R$)", "ÔNUS (R$)", "OBSERVAÇÕES"],
        [["Penhora Direta", "", "", "", ""],
         ["IDPJ", "", "", "", ""],
         ["Fraude à Execução", "", "", "", ""],
         ["Pauliana", "", "", "", ""]],
        [3.5, 2.6, 2.6, 2.6, 5.6],
        total_row=["TOTAL GERAL", "", "", "", ""],
    )
    _spacer(doc)

    # Visão geral dos atingíveis (3 tabelas)
    _sub_orange(doc, "VISÃO GERAL DOS ATINGÍVEIS")
    cols_at = ["ATIVO / MATRÍCULA", "TIPO DE ATIVO", "ÁREA / HA", "SIT. PRODUTIVA",
               "LIQUIDEZ", "% ATING.", "ÔNUS (R$)", "OBSERVAÇÕES"]
    ws_at = [2.6, 2.1, 1.6, 2.2, 1.9, 1.5, 2.0, 3.0]
    for tese in ["Penhora Direta", "IDPJ", "Fraude à Execução"]:
        _sub_gray(doc, tese)
        _grid_table(doc, cols_at, [], ws_at, min_rows=3)
        _spacer(doc, pts=2)

    # Principais considerações
    _sub_orange(doc, "PRINCIPAIS CONSIDERAÇÕES")
    cons = (dados.get("consideracoes_gerais") or "").strip()
    if cons:
        for para in cons.split("\n"):
            if para.strip():
                _body(doc, para.strip())
    else:
        _guidance(doc, "Registre aqui: visão geral da operação · principais riscos jurídicos e "
                       "patrimoniais · oportunidades identificadas · pontos críticos para aquisição · "
                       "impressões iniciais e contexto relevante do caso.")

    # ══ 1. VISÃO JURÍDICA ════════════════════════════════════════════════
    _sec_title(doc, "1. VISÃO JURÍDICA")

    creditos = dados.get("creditos") or []
    if not creditos:
        _body(doc, "Nenhum crédito identificado no relatório.")
    for idx, cred in enumerate(creditos, 1):
        cid = cred.get("id") or f"Crédito {idx}"
        _sub_orange(doc, f"1.{idx} {cid}")

        # Dados do Processo
        _sub_gray(doc, "Dados do Processo")
        _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], [
            ("Número do processo",            cred.get("numero_processo", "")),
            ("Vara / Comarca",                cred.get("vara_comarca", "")),
            ("Exequente",                     cred.get("exequente_info", "")),
            ("Executado(s)",                  cred.get("executados_info", "")),
            ("Data de distribuição",          cred.get("data_distribuicao", "")),
            ("SOP",                           cred.get("sop", "R$")),
            ("SAT",                           cred.get("sat", "R$")),
            ("Critério de atualização do SAT", cred.get("criterio_sat", "")),
            ("Honorários",                    cred.get("honorarios", "")),
            ("Lastro / Instrumento",          cred.get("lastro", "")),
            ("Data de Emissão",               cred.get("data_emissao", "")),
            ("Data do Vencimento",            cred.get("data_vencimento", "")),
            ("Assinaturas",                   cred.get("assinaturas", "")),
            ("Garantia",                      cred.get("garantia", "")),
            ("Status do Processo",            cred.get("status_processo", "")),
        ])
        _spacer(doc, pts=2)

        # Índices de Correção do Contrato
        _sub_gray(doc, "Índices de Correção do Contrato")
        _kv_table(doc, ["CRITÉRIO", "PARÂMETRO"], [
            ("Correção monetária",   cred.get("ind_cm", "")),
            ("Juros remuneratórios", cred.get("ind_jr", "")),
            ("Juros moratórios",     cred.get("ind_jm", "")),
            ("Multa moratória",      cred.get("ind_multa", "")),
            ("Capitalização",        cred.get("ind_cap", "")),
        ])
        _spacer(doc, pts=2)

        # Planilha Inicial
        _sub_gray(doc, "Planilha Inicial")
        _kv_table(doc, ["CRITÉRIO", "PARÂMETRO"], [
            ("Correção monetária",   cred.get("plan_cm", "")),
            ("Juros remuneratórios", cred.get("plan_jr", "")),
            ("Multa moratória",      cred.get("plan_multa", "")),
            ("Capitalização",        cred.get("plan_cap", "")),
            ("Ponderações",          cred.get("plan_ponderacoes", "")),
        ])
        _spacer(doc, pts=2)

        # Última Memória de Cálculo
        _sub_gray(doc, "Última Memória de Cálculo")
        _kv_table(doc, "DADOS DA MEMÓRIA DE CÁLCULO", [
            ("Data da Juntada",   cred.get("memoria_data_juntada", "")),
            ("Total Atualizado",  cred.get("memoria_total", "")),
            ("Data-base",         cred.get("memoria_data_base", "")),
            ("Índices aplicados", cred.get("memoria_indices", "")),
            ("Ponderações",       cred.get("memoria_ponderacoes", "")),
        ])
        _spacer(doc, pts=2)

        # Citação
        _sub_gray(doc, "Citação")
        citacoes = cred.get("citacoes") or []
        _grid_table(
            doc,
            ["EXECUTADO", "MODALIDADE", "DATA", "FLS."],
            [[c.get("executado", ""), c.get("modalidade", ""), c.get("data", ""), c.get("fls", "")]
             for c in citacoes],
            [5.5, 4.4, 3.5, 3.5],
            min_rows=1,
        )
        _spacer(doc, pts=2)

        # Embargos e/ou Exceção
        _sub_gray(doc, "Embargos e/ou Exceção")
        embargos = cred.get("embargos") or []
        if not embargos:
            embargos = [{}]
        for ei, emb in enumerate(embargos, 1):
            _para(doc, f"Embargo nº {ei}", bold=True, size=9.5, color=_TXT, before=3, after=2)
            _kv_label_table(doc, [
                ("Embargante",            emb.get("embargante", "")),
                ("Data da Distribuição",  emb.get("data_dist", "")),
                ("Tese dos Embargos",     emb.get("tese", "")),
                ("Principais Andamentos", emb.get("andamentos_resumo", "")),
                ("Status Atual",          emb.get("status", "")),
            ])
        _note(doc, "⊕ Replicar o quadro acima para cada embargo adicional identificado.")

        # Recursos
        _sub_gray(doc, "Recursos")
        recursos = cred.get("recursos") or []
        if not recursos:
            recursos = [{}]
        for ri, rec in enumerate(recursos, 1):
            _para(doc, f"Recurso nº {ri}", bold=True, size=9.5, color=_TXT, before=3, after=2)
            _kv_label_table(doc, [
                ("Recorrente",            rec.get("recorrente", "")),
                ("Decisão Recorrida",     rec.get("decisao_recorrida", "")),
                ("Data da Distribuição",  rec.get("data_dist", "")),
                ("Tese do Recurso",       rec.get("tese", "")),
                ("Principais Andamentos", rec.get("andamentos_resumo", "")),
                ("Status Atual",          rec.get("status", "")),
            ])
        _note(doc, "⊕ Replicar o quadro acima para cada recurso adicional identificado.")

        # Constrições Vigentes
        _sub_gray(doc, "Constrições Vigentes")
        constricoes = cred.get("constricoes") or []
        _grid_table(
            doc,
            ["TIPO", "DESCRIÇÃO", "VALOR (R$)", "STATUS"],
            [[c.get("tipo", ""), c.get("descricao", ""), c.get("valor", ""), c.get("status", "")]
             for c in constricoes],
            [4.0, 5.9, 3.0, 4.0],
            min_rows=1,
        )
        _note(doc, "Status possíveis: Ativa · Suspensa · Levantada · Contestada · Em discussão · "
                   "Aguardando avaliação · Aguardando expropriação")

        # Principais Andamentos Processuais
        _sub_gray(doc, "Principais Andamentos Processuais")
        andamentos = cred.get("andamentos") or []
        _grid_table(
            doc,
            ["DATA", "DESCRIÇÃO", "FLS. / EVENTO"],
            [[a.get("data", ""), a.get("descricao", ""), a.get("fls", "")] for a in andamentos],
            [2.6, 11.3, 3.0],
            min_rows=1,
        )
        _spacer(doc)

    # ══ 2. TESES DE RECUPERAÇÃO ══════════════════════════════════════════
    _sec_title(doc, "2. TESES DE RECUPERAÇÃO")

    cols_mat = ["Mat.", "Comarca", "Proprietário Atual", "Descrição do Imóvel", "Ônus Vigentes",
                "Fração", "VM (R$)", "VP (R$)", "Ônus Total (R$)", "Saldo (R$)"]
    ws_mat = [1.3, 1.7, 2.0, 2.5, 1.9, 1.2, 1.4, 1.4, 1.7, 1.4]

    # 2.1 Penhora Direta
    _sub_orange(doc, "2.1 Penhora Direta")
    _sub_gray(doc, "a. Ponderações e Observações Gerais")
    _guidance(doc, "Descrever aqui: situação geral dos imóveis localizados, gravames identificados, "
                   "bloqueios de matrícula, imóveis não localizados, pesquisas pendentes e demais "
                   "observações relevantes.")
    _sub_gray(doc, "b. Matrículas Mapeadas")
    _grid_table(doc, cols_mat, [], ws_mat, total_row=["TOTAL", "", "", "", "", "", "", "", "", ""], min_rows=2)
    _spacer(doc)

    # 2.2 IDPJ
    _sub_orange(doc, "2.2 IDPJ")
    _sub_gray(doc, "a. Resumo da Tese")
    _guidance(doc, "Fundamento principal · Histórico dos fatos · Objetivo do IDPJ · Principais elementos "
                   "probatórios identificados (procurações, representação conjunta, declarações em processo, "
                   "confusão patrimonial, etc.).")
    _sub_gray(doc, "b. Empresa-Alvo")
    _kv_table(doc, ["CAMPO", "DADO"], [
        ("Razão social da empresa-alvo",        ""),
        ("CNPJ",                                ""),
        ("Fundamentação da tese",               "Simulação / Desvio de finalidade / Confusão patrimonial"),
        ("Crédito mais adequado para propositura", ""),
    ])
    _spacer(doc, pts=2)
    _sub_gray(doc, "c. Cronologia Societária — Atos Relevantes")
    _grid_table(doc, ["DATA", "ATO", "DETALHAMENTO"], [], [2.6, 4.5, 9.8], min_rows=2)
    _spacer(doc, pts=2)
    _sub_gray(doc, "d. Evidências de Controle Informal")
    _guidance(doc, "Espaço destinado à consolidação dos principais elementos probatórios que sustentam a "
                   "tese: fatos relevantes identificados durante a investigação, documentos, prints, imagens, "
                   "pesquisas patrimoniais, alterações societárias, procurações, manifestações processuais e "
                   "demais evidências que demonstrem a viabilidade da medida.")
    _sub_gray(doc, "e. Ativos Atingíveis via IDPJ  —  ~R$ [___] (VM) / ~R$ [___] (VP)")
    _grid_table(doc, cols_mat, [], ws_mat, total_row=["TOTAL", "", "", "", "", "", "", "", "", ""], min_rows=2)
    _spacer(doc, pts=2)
    _sub_gray(doc, "f. Upside Identificado (se houver)")
    _guidance(doc, "Descrever eventual upside — ex: recebíveis futuros de compra e venda em andamento, "
                   "participações societárias, créditos a receber, outros ativos. Indicar valor estimado e "
                   "prazo de realização.")
    _spacer(doc)

    # 2.3 Fraude à Execução
    _sub_orange(doc, "2.3 Fraude à Execução")
    _sub_gray(doc, "a. Resumo da Tese")
    _guidance(doc, "Descrever os negócios jurídicos impugnáveis (doações, transferências, usufruto) "
                   "praticados no curso ou após o ajuizamento da execução. Indicar a data do ajuizamento e a "
                   "data de cada ato para enquadramento da tese. Identificar qual crédito é mais adequado "
                   "para a propositura.")
    _sub_gray(doc, "b. Má-fé e Insolvência")
    _guidance(doc, "Registrar os elementos que demonstram a má-fé do adquirente (consilium fraudis) e a "
                   "insolvência do devedor decorrente do ato — pressupostos da fraude à execução / ação pauliana.")
    _sub_gray(doc, "c. Ativos Atingíveis via Fraude à Execução  —  ~R$ [___] (VM) / ~R$ [___] (VP)")
    _grid_table(doc, cols_mat, [], ws_mat, total_row=["TOTAL", "", "", "", "", "", "", "", "", ""], min_rows=2)
    _spacer(doc)

    # 2.4 Outras teses
    _sub_orange(doc, "2.4 Outras teses")
    _guidance(doc, "Registrar outras teses de recuperação eventualmente aplicáveis ao caso.")
    _spacer(doc)

    # ══ 3. PASSIVO ═══════════════════════════════════════════════════════
    _sec_title(doc, "3. PASSIVO")
    _note(doc, "Mapear o passivo integral dos devedores e das empresas-alvo. O passivo é fator "
               "determinante na avaliação do risco real de recuperação.")

    # 3.1 Passivo Fiscal
    _sub_orange(doc, "3.1 Passivo Fiscal")
    _sub_gray(doc, "e-CAC")
    _grid_table(doc, ["NOME", "CPF / CNPJ", "SALDO e-CAC (R$)"], [], [7.0, 5.0, 4.9], min_rows=2)
    _spacer(doc, pts=2)
    _sub_gray(doc, "Execuções Fiscais")
    cols_exec = ["Nº CNJ", "Vinculado a", "Distribuição", "Exequente", "Valor Causa (R$)", "SAT Est. (R$)", "Status"]
    ws_exec = [3.0, 2.2, 2.0, 2.8, 2.3, 2.3, 2.3]
    _grid_table(doc, cols_exec, [], ws_exec, min_rows=2)
    _spacer(doc)

    # 3.2 Passivo Trabalhista
    _sub_orange(doc, "3.2 Passivo Trabalhista")
    cols_trab = ["Nº CNJ", "Vinculado a", "Distribuição", "Autor", "Valor Causa (R$)", "SAT Est. (R$)", "Status"]
    _grid_table(doc, cols_trab, [], ws_exec, min_rows=2)
    _spacer(doc)

    # 3.3 Passivo Cível (Terceiros)
    _sub_orange(doc, "3.3 Passivo Cível (Terceiros)")
    cols_civ = ["Nº CNJ", "Vinculado a", "Distribuição", "Exequente", "Valor Causa (R$)", "SAT Est. (R$)", "Obs."]
    _grid_table(doc, cols_civ, [], ws_exec, min_rows=2)
    _spacer(doc, pts=2)

    # Principais Credores
    _sub_gray(doc, "Principais Credores")
    _grid_table(
        doc,
        ["CREDOR", "NATUREZA DO CRÉDITO", "VALOR ESTIMADO (R$)", "GARANTIA / PREFERÊNCIA", "STATUS", "OBSERVAÇÕES"],
        [], [3.0, 2.9, 2.5, 3.0, 2.1, 3.4], min_rows=2,
    )
    _note(doc, "Pontos de atenção sobre o passivo: hipotecas preferentes, créditos arrematados por "
               "terceiros, passivo oculto potencial, ações de reintegração, litígios relevantes, etc.")
    _spacer(doc)

    # ══ 5. PENDÊNCIAS ════════════════════════════════════════════════════
    _sec_title(doc, "5. PENDÊNCIAS")
    _note(doc, "Campo destinado ao registro de pendências relacionadas à pesquisa de bens, obtenção de "
               "escrituras, avaliações de imóveis e demais diligências ainda necessárias para a conclusão "
               "da análise do caso.")
    _grid_table(doc, ["PENDÊNCIA", "PROTOCOLO", "PRAZO / STATUS"], [], [8.9, 4.0, 4.0], min_rows=2)
    _spacer(doc)

    # ══ 6. ELABORAÇÃO E REVISÃO ══════════════════════════════════════════
    _sec_title(doc, "6. ELABORAÇÃO E REVISÃO")
    _grid_table(
        doc,
        ["FUNÇÃO", "NOME", "DATA"],
        [["Advogada(o) Responsável pela Análise", "", ""],
         ["Advogada(o) Revisor", "", ""]],
        [8.4, 5.5, 3.0],
    )
    _spacer(doc)

    # Rodapé de confidencialidade (corpo)
    _para(doc,
          "Documento confidencial. Uso interno exclusivo do Time PPA — Invista. "
          "Vedada a reprodução ou distribuição a terceiros sem autorização prévia.",
          size=8, color=_TXT_MUTE, italic=True, before=8, after=0,
          align=WD_ALIGN_PARAGRAPH.CENTER)

    # ── Salvar ────────────────────────────────────────────────────────────
    nome = re.sub(r"[^\w\s-]", "", dados.get("nome_caso", "") or "caso").strip().replace(" ", "_")[:40] or "caso"
    caminho = os.path.join(tempfile.gettempdir(), f"Dossie_PPA_{nome}.docx")
    doc.save(caminho)
    return caminho


# ══════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════

def gerar_dossie_word(relatorio: str, client, model: str) -> str:
    """Extrai dados do relatório via Gemini e gera o Word no formato PPA v3.0."""
    dados = _extrair_dados(relatorio, client, model)
    return _build_doc(dados)
