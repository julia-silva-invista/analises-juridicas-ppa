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
from docx.table import Table as _DocxTable
from docx.text.paragraph import Paragraph as _DocxParagraph

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

# Siglas que devem conservar a grafia técnica mesmo quando o restante do
# texto vier integralmente em caixa alta.
_SIGLAS_PRESERVADAS = {
    "AR", "CAC", "CCB", "CNJ", "CNPJ", "CPF", "ID", "IDPJ", "INPC", "IPCA",
    "OAB", "OJ", "PPA", "RG", "SA", "SAT", "SOP", "UF", "VM", "VP",
}
_PARTICULAS_NOME = {
    "a", "as", "com", "da", "das", "de", "do", "dos", "e", "em", "na",
    "nas", "no", "nos", "para", "por",
}
_RE_PALAVRA = re.compile(
    r"[A-Za-zÀ-ÖØ-öø-ÿ]+(?:[.'’][A-Za-zÀ-ÖØ-öø-ÿ]+)*\.?"
)


def _normalizar_caixa_alta(texto) -> str:
    """
    Converte texto predominantemente em caixa alta para capitalização natural.

    Nomes como ``JULIA DE OLIVEIRA`` tornam-se ``Julia de Oliveira``. Siglas
    jurídicas e financeiras conhecidas permanecem em caixa alta. Textos que
    já possuem capitalização normal não são alterados.
    """
    valor = str(texto if texto is not None else "")
    if valor.strip() == "e-CAC":
        return valor
    letras = [c for c in valor if c.isalpha()]
    if len(letras) < 2:
        return valor
    proporcao_maiusculas = sum(c.isupper() for c in letras) / len(letras)
    if proporcao_maiusculas < 0.72:
        return valor

    indice = 0

    def _ajustar(match):
        nonlocal indice
        palavra = match.group(0)
        chave = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿ]", "", palavra).upper()
        minuscula = palavra.lower()
        if chave in _SIGLAS_PRESERVADAS:
            resultado = palavra.upper()
        elif indice > 0 and minuscula.rstrip(".") in _PARTICULAS_NOME:
            resultado = minuscula
        else:
            resultado = minuscula[:1].upper() + minuscula[1:]
        indice += 1
        return resultado

    return _RE_PALAVRA.sub(_ajustar, valor)


def _normalizar_dados(valor):
    """Aplica a capitalização natural recursivamente aos dados extraídos."""
    if isinstance(valor, dict):
        return {chave: _normalizar_dados(item) for chave, item in valor.items()}
    if isinstance(valor, list):
        return [_normalizar_dados(item) for item in valor]
    if isinstance(valor, str):
        return _normalizar_caixa_alta(valor)
    return valor


def _alinhamento_conteudo(texto):
    """Justifica conteúdo narrativo; mantém dados curtos alinhados à esquerda."""
    linhas = str(texto if texto is not None else "").splitlines() or [""]
    if max(len(linha.strip()) for linha in linhas) >= 80:
        return WD_ALIGN_PARAGRAPH.JUSTIFY
    return None


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


def _nao_dividir_linha(row):
    """Evita que uma linha de tabela seja fragmentada entre duas páginas."""
    trp = row._tr.get_or_add_trPr()
    if trp.find(qn("w:cantSplit")) is None:
        trp.append(OxmlElement("w:cantSplit"))


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
    if align is not None:
        p.alignment = align
    parts = str(text if text is not None else "").split("\n")
    for j, line in enumerate(parts):
        r = p.add_run(_normalizar_caixa_alta(line))
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
    _nao_dividir_linha(row)
    for i, lab in enumerate(labels):
        c = row.cells[i]
        _set_bg(c, _LARANJA); _set_borders(c); _cell_pad(c, left=55, right=55)
        _write(c, lab, bold=True, size=8.2, color=_BRANCO)
    if ws:
        _widths(row, ws)
    return row


def _span_header(table, text):
    """Header laranja que abrange todas as colunas (ex.: DADOS DA MEMÓRIA DE CÁLCULO)."""
    row = table.add_row()
    _nao_dividir_linha(row)
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
    t.autofit = False
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
        _write(cv, value, color=_TXT, align=_alinhamento_conteudo(value))
        _widths(r, [w_label, w_value])
    return t


def _kv_label_table(doc, rows, w_label=5.9, w_value=11.0):
    """Tabela 2 col com coluna-label cinza (embargo/recurso), sem header laranja."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.autofit = False
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for label, value in rows:
        r = t.add_row()
        cl, cv = r.cells[0], r.cells[1]
        _set_bg(cl, _CINZA); _set_borders(cl); _cell_pad(cl)
        _write(cl, label, bold=True, color=_TXT)
        _set_bg(cv, _BRANCO); _set_borders(cv); _cell_pad(cv)
        _write(cv, value, color=_TXT, align=_alinhamento_conteudo(value))
        _widths(r, [w_label, w_value])
    return t


def _grid_table(doc, headers, rows, ws, total_row=None, min_rows=1):
    """Tabela multi-coluna: header laranja + linhas brancas (+ TOTAL cinza opcional)."""
    t = doc.add_table(rows=0, cols=len(headers))
    t.style = "Table Grid"
    t.autofit = False
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
            _write(c, val, align=_alinhamento_conteudo(val))
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


def _cards(valor):
    """Converte texto, lista ou dicionário em cartões narrativos homogêneos."""
    if not valor:
        return []
    itens = valor if isinstance(valor, list) else [valor]
    resultado = []
    for item in itens:
        if isinstance(item, dict):
            texto = (
                item.get("texto")
                or item.get("descricao")
                or item.get("detalhamento")
                or item.get("analise")
                or ""
            )
            titulo = item.get("titulo") or item.get("fato") or item.get("ato") or ""
            referencia = item.get("referencia") or item.get("fonte") or ""
            if str(texto).strip() or str(titulo).strip():
                resultado.append({
                    "titulo": str(titulo).strip(),
                    "texto": str(texto).strip(),
                    "referencia": str(referencia).strip(),
                })
            continue
        for trecho in re.split(r"\n\s*\n|\n(?=\S)", str(item)):
            if trecho.strip():
                resultado.append({"titulo": "", "texto": trecho.strip(), "referencia": ""})
    return resultado


def _write_analysis_cell(cell, texto, titulo="", referencia=""):
    """Escreve narrativa jurídica justificada e referência destacada na célula."""
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p = cell.paragraphs[0]
    if titulo:
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        p.paragraph_format.keep_with_next = True
        rtitle = p.add_run(_normalizar_caixa_alta(titulo))
        _apply_font(rtitle, bold=True, size=8.8, color=_LARANJA, italic=False)
        p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.08
    r = p.add_run(_normalizar_caixa_alta(texto))
    _apply_font(r, bold=False, size=9.3, color=_TXT, italic=False)
    if referencia:
        pref = cell.add_paragraph()
        pref.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pref.paragraph_format.space_before = Pt(4)
        pref.paragraph_format.space_after = Pt(0)
        rr = pref.add_run(_normalizar_caixa_alta(referencia))
        _apply_font(rr, bold=False, size=7.8, color=_TXT_MUTE, italic=True)


def _analysis_box(doc, texto, titulo="", referencia=""):
    """Quadro formal para uma unidade autônoma de análise jurídica."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.autofit = False
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    b = t.rows[0]
    if len(str(texto)) <= 900:
        _nao_dividir_linha(b)
    bc = b.cells[0]
    _set_bg(bc, "FFFDFC")
    _set_borders(bc, color="E8B9A8", sz=6)
    _cell_pad(bc, top=115, left=135, bottom=115, right=135)
    _write_analysis_cell(bc, texto, titulo=titulo, referencia=referencia)
    bc.width = Cm(16.9)
    return t


def _analysis_cards(doc, valor, placeholder=""):
    """Renderiza a análise em quadros separados, um fato ou fundamento por cartão."""
    itens = _cards(valor)
    if not itens:
        if placeholder:
            _guidance(doc, placeholder)
        return
    for i, item in enumerate(itens):
        _analysis_box(
            doc,
            item.get("texto", ""),
            titulo=item.get("titulo", ""),
            referencia=item.get("referencia", ""),
        )
        if i < len(itens) - 1:
            _spacer(doc, pts=2)


def _asset_cards(doc, ativos, tese=None, placeholder=True):
    """Exibe ativos em fichas legíveis, sem comprimir narrativa em dez colunas."""
    registros = [
        a for a in (ativos or [])
        if not tese or tese.casefold() in str(a.get("tese", "")).strip().casefold()
    ]
    if not registros:
        if placeholder:
            _kv_table(doc, ["ATIVO", "INFORMAÇÃO"], [
                ("Matrícula / identificação", ""),
                ("Proprietário atual", ""),
                ("Descrição e situação jurídica", ""),
                ("Avaliação e potencial de recuperação", ""),
            ], w_label=5.0, w_value=11.9)
        return
    for i, ativo in enumerate(registros):
        identificacao = ativo.get("matricula") or ativo.get("ativo") or f"Ativo nº {i + 1}"
        comarca = ativo.get("comarca", "")
        titulo = " · ".join(x for x in [identificacao, comarca] if x)
        descricao = " · ".join(x for x in [
            ativo.get("tipo_ativo", ""),
            ativo.get("descricao", ""),
            ativo.get("area", ""),
            ativo.get("situacao_produtiva", ""),
        ] if x)
        situacao = " · ".join(x for x in [
            ativo.get("onus_vigentes", ""),
            f"Fração atingível: {ativo.get('fracao_atingivel', '')}" if ativo.get("fracao_atingivel") else "",
            f"Liquidez: {ativo.get('liquidez', '')}" if ativo.get("liquidez") else "",
        ] if x)
        valores = " · ".join(x for x in [
            f"VM: {ativo.get('vm', '')}" if ativo.get("vm") else "",
            f"VP: {ativo.get('vp', '')}" if ativo.get("vp") else "",
            f"Ônus: {ativo.get('onus_total', '')}" if ativo.get("onus_total") else "",
            f"Saldo estimado: {ativo.get('saldo', '')}" if ativo.get("saldo") else "",
        ] if x)
        _kv_table(doc, titulo.upper(), [
            ("Proprietário atual", ativo.get("proprietario_atual", "")),
            ("Descrição do ativo", descricao),
            ("Situação jurídica", situacao),
            ("Avaliação", valores),
            ("Observações", ativo.get("observacoes", "")),
        ], w_label=4.3, w_value=12.6)
        if i < len(registros) - 1:
            _spacer(doc, pts=3)


def _process_cards(doc, registros, parte_label):
    """Quadros processuais do passivo, com situação em largura integral."""
    t = doc.add_table(rows=0, cols=2)
    t.style = "Table Grid"
    t.autofit = False
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _fill_process_cards(t, registros, parte_label)
    return t


def _fill_process_cards(table, registros, parte_label):
    """Preenche uma tabela de duas colunas com um ou mais processos."""
    for row in list(table.rows):
        table._tbl.remove(row._tr)
    dados = list(registros or [])
    if not dados:
        _span_header(table, "DADOS PROCESSUAIS")
        r = table.add_row()
        _set_bg(r.cells[0], _CINZA); _set_borders(r.cells[0]); _cell_pad(r.cells[0])
        _write(r.cells[0], "Processo", bold=True)
        _set_bg(r.cells[1], _BRANCO); _set_borders(r.cells[1]); _cell_pad(r.cells[1])
        _write(r.cells[1], "")
        _widths(r, [4.3, 12.6])
        return
    for registro in dados:
        if isinstance(registro, dict):
            numero = registro.get("numero_cnj") or registro.get("numero") or ""
            vinculado = registro.get("vinculado_a") or registro.get("vinculado") or ""
            distribuicao = registro.get("distribuicao") or ""
            parte = registro.get("parte_contraria") or registro.get("exequente") or registro.get("autor") or ""
            valor = registro.get("valor_causa") or ""
            sat = registro.get("sat_estimado") or registro.get("sat") or ""
            status = registro.get("status") or registro.get("situacao") or ""
        else:
            valores = list(registro) + [""] * 7
            numero, vinculado, distribuicao, parte, valor, sat, status = valores[:7]
        _span_header(table, f"Processo nº {numero}" if numero else "Processo")
        for label, value in [
            ("Vinculado a", vinculado),
            ("Distribuição", distribuicao),
            (parte_label, parte),
            ("Valor da causa", valor),
            ("SAT estimado", sat),
            ("Situação processual", status),
        ]:
            r = table.add_row()
            if len(str(value)) <= 900:
                _nao_dividir_linha(r)
            _set_bg(r.cells[0], _CINZA); _set_borders(r.cells[0]); _cell_pad(r.cells[0])
            _write(r.cells[0], label, bold=True, size=8.8)
            _set_bg(r.cells[1], _BRANCO); _set_borders(r.cells[1]); _cell_pad(r.cells[1], top=70, bottom=70)
            _write(r.cells[1], value, size=8.8, align=_alinhamento_conteudo(value), valign=False)
            _widths(r, [4.3, 12.6])


# ══════════════════════════════════════════════════════════════════════════
# Helpers de parágrafo (títulos, notas, texto)
# ══════════════════════════════════════════════════════════════════════════

def _para(doc, text, bold=False, size=9.5, color=_TXT, italic=False,
          before=0, after=4, align=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    if align is not None:
        p.alignment = align
    r = p.add_run(_normalizar_caixa_alta(text))
    _apply_font(r, bold, size, color, italic)
    return p


def _sec_title(doc, text):
    p = _para(doc, text, bold=True, size=13, color=_LARANJA, before=14, after=5)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    return p


def _sub_orange(doc, text):
    p = _para(doc, text, bold=True, size=11, color=_LARANJA, before=10, after=4)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    return p


def _sub_gray(doc, text):
    p = _para(doc, text, bold=True, size=10, color=_TXT, before=7, after=3)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
    return p


def _note(doc, text):
    return _para(
        doc, text, size=8, color=_TXT_MUTE, italic=True, before=2, after=6,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    )


def _body(doc, text):
    return _para(
        doc, text, size=9.5, color=_TXT, before=0, after=5,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    )


def _guidance(doc, text):
    """Caixa de orientação com fundo FEF0EB (célula única)."""
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = t.rows[0].cells[0]
    _nao_dividir_linha(t.rows[0])
    _set_bg(c, _DESTAQUE); _set_borders(c, color="F3D9CF"); _cell_pad(c, top=90, bottom=90)
    _write(
        c, text, size=9, color=_TXT, italic=True,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    )
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
Com base no material jurídico abaixo (extração completa do(s) processo(s) e, quando houver, relatório
consolidado de apoio), extraia os dados para preencher um dossiê no formato JSON.
Extraia APENAS o que está explicitamente no material. Use string vazia "" para o que não constar.
NÃO invente números, nomes ou datas. Responda SOMENTE com o JSON, sem texto adicional.

═══ REGRA 1 — REFERÊNCIA OBRIGATÓRIA DA FONTE ═══
Para CADA informação preenchida, inclua ao final, entre parênteses, a referência da fonte no processo —
use o parâmetro que o tribunal daquela execução usa (fls., Mov., Evento, ID ou página).
Ex.: "R$ 12.450.000,00 (fls. 45)", "IPCA (cláusula 4ª — fls. 12)", "Penhora deferida (Mov. 120)".
Se a referência realmente não constar, deixe o valor sem os parênteses (não invente a referência).

═══ REGRA 2 — CAMPOS QUE DEVEM SER PREENCHIDOS ═══
- Índices de Correção do Contrato (ind_*): busque no título executivo (CCB/contrato/duplicata) e na
  petição inicial — correção monetária, juros remuneratórios, juros moratórios, multa, capitalização.
- Planilha Inicial (plan_*): a memória de cálculo que instruiu a petição inicial.
- Última Memória de Cálculo (memoria_*): a memória de débito MAIS RECENTE juntada aos autos
  (data da juntada, total atualizado, data-base, índices aplicados, ponderações).
- Liste TODOS os embargos à execução, TODAS as exceções/objeções de pré-executividade e TODOS os
  recursos que aparecerem (não apenas o primeiro). No campo "tipo" de cada defesa informe
  "Embargos à Execução" ou "Exceção de Pré-Executividade".

═══ REGRA 3 — CITAÇÃO (uma entrada por executado) ═══
Gere UMA entrada por executado da execução. Se o executado FOI citado, informe modalidade (AR/OJ/
Edital/Precatória), data e fls. da citação. Se NÃO foi citado, ponha modalidade="Pendente" e liste no
campo "data" as datas das tentativas e no campo "fls" as páginas correspondentes (na mesma linha).

═══ REGRA 4 — PENHORAS E ANDAMENTOS ═══
- Constrições: no campo "status", se a penhora/constrição foi DEFERIDA, inclua a DATA da decisão de
  deferimento. Ex.: "Penhora deferida em 15/03/2024 (fls. 120)", "Sisbajud deferido em 02/2024 (Mov. 88)".
- andamentos: liste TODOS os andamentos processuais que constarem no material E no relatório de apoio —
  NÃO resuma, NÃO omita e NÃO limite a quantidade. Cada andamento com data, descrição objetiva e
  referência (Mov./fls./Evento/ID).

═══ REGRA 5 — FORMATAÇÃO DO TEXTO ═══
- Nunca escreva nomes de pessoas físicas ou jurídicas integralmente em caixa alta.
- Use capitalização natural: "Julia de Oliveira Bernardo da Silva", "Empresa Agrícola Ltda.".
- Preserve em caixa alta somente siglas técnicas, como CPF, CNPJ, CNJ, OAB, SAT, SOP, IDPJ, VM e VP.
- Redija títulos, descrições, teses, status e andamentos com maiúsculas e minúsculas normais.

═══ REGRA 6 — ANÁLISE JURÍDICA E QUADROS ═══
- Preencha também as teses de recuperação, os ativos, o passivo e as pendências com tudo o que constar
  no relatório consolidado. Não deixe essas seções como modelo vazio se o relatório trouxer a informação.
- Em campos de análise, produza cartões autônomos: um fato, fundamento ou conclusão por item.
- Cada cartão deve ter título objetivo, texto jurídico completo e referência da fonte separada.
- Preserve a cronologia, os requisitos legais, os riscos, as ressalvas e a conclusão. Não simplifique
  a ponto de eliminar fatos relevantes e não repita o mesmo conteúdo em cartões diferentes.
- Use linguagem formal, assertiva e tecnicamente precisa. Diferencie fato comprovado, indício,
  inferência jurídica, risco e diligência pendente.

{
  "nome_caso": "identificador curto (ex: BASF x São Lourenço)",
  "data_analise": "DD/MM/AAAA ou vazio",
  "exequentes": "Nome · CNPJ ... (fls.)",
  "executados": "Nome1 · CPF/CNPJ ...; Nome2 ... (fls.)",
  "sat_total": "R$ ... (fls.)",
  "passivo_fiscal": "", "passivo_trabalhista": "", "passivo_civel": "",
  "risco_juridico": "resumo dos principais riscos (com refs)",
  "consideracoes_gerais": "2-4 parágrafos de análise geral (separe parágrafos com \\n)",
  "teses_principais": "síntese das teses aplicáveis",
  "total_atingivel_vm": "R$ ...",
  "total_atingivel_vp": "R$ ...",
  "visao_consolidada_ativos": [
    {"tese": "Penhora Direta | IDPJ | Fraude à Execução | outra", "vm": "R$ ...", "vp": "R$ ...", "onus": "R$ ...", "observacoes": ""}
  ],
  "ativos": [
    {
      "tese": "Penhora Direta | IDPJ | Fraude à Execução | outra",
      "matricula": "Matrícula nº ...", "comarca": "Cidade/UF", "proprietario_atual": "",
      "tipo_ativo": "apartamento, galpão, terreno etc.", "descricao": "", "area": "",
      "situacao_produtiva": "", "liquidez": "", "fracao_atingivel": "",
      "onus_vigentes": "", "vm": "R$ ...", "vp": "R$ ...", "onus_total": "R$ ...",
      "saldo": "R$ ...", "observacoes": ""
    }
  ],
  "teses_recuperacao": {
    "penhora_direta": {
      "analise": [{"titulo": "", "texto": "fato, fundamento e conclusão", "referencia": "fls./Mov./ID/Evento"}]
    },
    "idpj": {
      "resumo": [{"titulo": "", "texto": "", "referencia": ""}],
      "empresa_alvo": {
        "razao_social": "", "cnpj": "", "cnae_principal": "", "atividades_secundarias": "",
        "socio_atual": "", "endereco_fiscal": "", "fundamentacao": "",
        "credito_propositura": ""
      },
      "cronologia": [{"data": "", "ato": "", "detalhamento": "", "referencia": ""}],
      "evidencias": [{"titulo": "", "texto": "", "referencia": ""}],
      "upside": [{"titulo": "", "texto": "", "referencia": ""}]
    },
    "fraude_execucao": {
      "resumo": [{"titulo": "", "texto": "", "referencia": ""}],
      "ma_fe_insolvencia": [{"titulo": "", "texto": "", "referencia": ""}]
    },
    "outras": [{"titulo": "", "texto": "", "referencia": ""}]
  },
  "passivo": {
    "e_cac": [{"nome": "", "cpf_cnpj": "", "saldo": ""}],
    "fiscais": [{"numero_cnj": "", "vinculado_a": "", "distribuicao": "", "parte_contraria": "", "valor_causa": "", "sat_estimado": "", "status": ""}],
    "trabalhistas": [{"numero_cnj": "", "vinculado_a": "", "distribuicao": "", "parte_contraria": "", "valor_causa": "", "sat_estimado": "", "status": ""}],
    "civeis": [{"numero_cnj": "", "vinculado_a": "", "distribuicao": "", "parte_contraria": "", "valor_causa": "", "sat_estimado": "", "status": ""}],
    "principais_credores": [{"credor": "", "natureza": "", "valor": "", "garantia_preferencia": "", "status": "", "observacoes": ""}],
    "pontos_atencao": [{"titulo": "", "texto": "", "referencia": ""}]
  },
  "pendencias": [{"pendencia": "", "protocolo": "", "prazo_status": ""}],
  "creditos": [
    {
      "id": "Crédito [Credor]",
      "numero_processo": "0000000-00.0000.0.00.0000",
      "vara_comarca": "1ª Vara Cível — Cidade/UF",
      "exequente_info": "Nome · Adv: Escritório | OAB",
      "executados_info": "Nome · Adv: Escritório | OAB",
      "data_distribuicao": "DD/MM/AAAA (fls.)",
      "sop": "R$ ... (fls.)", "sat": "R$ ... (fls.)", "criterio_sat": "ex: INPC + 12% a.a JM (fls.)",
      "honorarios": "ex: 10% (fls.)",
      "lastro": "CCB nº / Contrato nº (fls.)", "data_emissao": "DD/MM/AAAA (fls.)", "data_vencimento": "DD/MM/AAAA (fls.)",
      "assinaturas": "Nomes (fls.)", "garantia": "descrição da garantia (fls.)",
      "status_processo": "ex: em fase de penhora (Mov.)",
      "ind_cm": "correção monetária (fls.)", "ind_jr": "juros remuneratórios (fls.)", "ind_jm": "juros moratórios (fls.)", "ind_multa": "multa (fls.)", "ind_cap": "capitalização (fls.)",
      "plan_cm": "(fls.)", "plan_jr": "(fls.)", "plan_multa": "(fls.)", "plan_cap": "(fls.)", "plan_ponderacoes": "(fls.)",
      "memoria_data_juntada": "DD/MM/AAAA (fls.)", "memoria_total": "R$ ... (fls.)", "memoria_data_base": "DD/MM/AAAA",
      "memoria_indices": "índices aplicados (fls.)", "memoria_ponderacoes": "",
      "citacoes": [{"executado": "Nome", "modalidade": "AR/OJ/Edital ou Pendente", "data": "data citação OU datas das tentativas", "fls": "fls. citação OU fls. das tentativas"}],
      "embargos": [{"tipo": "Embargos à Execução | Exceção de Pré-Executividade", "embargante": "", "data_dist": "(fls.)", "tese": "", "andamentos_resumo": "(com refs)", "status": ""}],
      "recursos": [{"recorrente": "", "decisao_recorrida": "", "data_dist": "(fls.)", "tese": "", "andamentos_resumo": "(com refs)", "status": ""}],
      "constricoes": [{"tipo": "", "descricao": "(fls.)", "valor": "R$ ... (fls.)", "status": "Ativa | Deferida em DD/MM/AAAA (fls.)"}],
      "andamentos": [{"data": "DD/MM/AAAA", "descricao": "descrição objetiva", "fls": "Mov./fls./Evento/ID"}]
    }
  ]
}

MATERIAL:
"""


def _extrair_dados(texto_completo: str, relatorio: str, client, model: str) -> dict:
    primario = (texto_completo or "").strip() or (relatorio or "").strip()
    prompt = _PROMPT_DOSSIE + primario[:900_000]
    rel = (relatorio or "").strip()
    if rel and rel != primario:
        prompt += "\n\n═══ RELATÓRIO CONSOLIDADO (resumo estruturado de apoio) ═══\n" + rel[:40_000]
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


_PROMPT_COMPLETAR = """\
Abaixo está um dossiê em JSON (parcialmente preenchido) e o RELATÓRIO consolidado do caso.
Confira, campo a campo, quais estão VAZIOS ("") e preencha-os com informação que CONSTE no RELATÓRIO.
Regras:
- NÃO altere campos que já têm valor. NÃO invente — só preencha o que constar no relatório.
- Mantenha EXATAMENTE a mesma estrutura, chaves e lista de créditos do JSON atual.
- Inclua a referência da fonte (fls./Mov./ID/Evento) quando o relatório trouxer.
- Se um andamento/embargo/recurso constar no relatório e faltar no JSON, ACRESCENTE à lista.
Responda SOMENTE com o JSON completo.

JSON ATUAL:
{json_atual}

RELATÓRIO:
{relatorio}
"""


def _completar_com_relatorio(dados: dict, relatorio: str, client, model: str) -> dict:
    """Passada final: preenche campos vazios do dossiê com o que constar no relatório."""
    rel = (relatorio or "").strip()
    if not dados or not dados.get("creditos") or not rel:
        return dados
    try:
        prompt = _PROMPT_COMPLETAR.format(
            json_atual=json.dumps(dados, ensure_ascii=False)[:150_000],
            relatorio=rel[:150_000],
        )
        config = types.GenerateContentConfig(response_mime_type="application/json")

        def _fn():
            return client.models.generate_content(
                model=model,
                contents=[types.Content(role="user", parts=[types.Part(text=prompt)])],
                config=config,
            ).text

        raw = _retry(_fn, tentativas=2, espera_base=10)
        raw = re.sub(r"^```[a-z]*\n?", "", raw.strip(), flags=re.IGNORECASE)
        raw = re.sub(r"\n?```$", "", raw.strip())
        completo = json.loads(raw)
        # Só aceita se preservar (ou ampliar) a lista de créditos — evita perder dados
        if isinstance(completo, dict) and len(completo.get("creditos") or []) >= len(dados.get("creditos") or []):
            return completo
        return dados
    except Exception:
        return dados


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
    dados = _normalizar_dados(dados or {})
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
        ("Total atingível mapeado — VM", dados.get("total_atingivel_vm", "") or "R$ [___]"),
        ("Total atingível mapeado — VP", dados.get("total_atingivel_vp", "") or "R$ [___]"),
        ("Passivo total identificado",   passivo_cell),
        ("Tese(s) principal(is)",        dados.get("teses_principais", "")),
        ("Risco Jurídico",               dados.get("risco_juridico", "")),
    ])
    _spacer(doc)

    # Visão consolidada dos ativos
    _sub_orange(doc, "VISÃO CONSOLIDADA DOS ATIVOS")
    _note(doc, "Quadro-resumo dos ativos atingíveis por tese. Os valores já descontam os ônus incidentes sobre os ativos.")
    resumo_ativos = dados.get("visao_consolidada_ativos") or []
    linhas_resumo = [[
        item.get("tese", ""),
        item.get("vm", ""),
        item.get("vp", ""),
        item.get("onus", ""),
        item.get("observacoes", ""),
    ] for item in resumo_ativos]
    if not linhas_resumo:
        linhas_resumo = [
            ["Penhora Direta", "", "", "", ""],
            ["IDPJ", "", "", "", ""],
            ["Fraude à Execução", "", "", "", ""],
            ["Pauliana", "", "", "", ""],
        ]
    _grid_table(
        doc,
        ["TESE", "VM (R$)", "VP (R$)", "ÔNUS (R$)", "OBSERVAÇÕES"],
        linhas_resumo,
        [3.5, 2.6, 2.6, 2.6, 5.6],
        total_row=[
            "TOTAL GERAL",
            dados.get("total_atingivel_vm", ""),
            dados.get("total_atingivel_vp", ""),
            "",
            "",
        ],
    )
    _spacer(doc)

    # Visão geral dos atingíveis em fichas — narrativa não é comprimida em colunas estreitas.
    _sub_orange(doc, "VISÃO GERAL DOS ATINGÍVEIS")
    ativos = dados.get("ativos") or []
    for tese in ["Penhora Direta", "IDPJ", "Fraude à Execução"]:
        _sub_gray(doc, tese)
        _asset_cards(doc, ativos, tese=tese, placeholder=True)
        _spacer(doc, pts=2)

    # Principais considerações
    _sub_orange(doc, "PRINCIPAIS CONSIDERAÇÕES")
    cons = (dados.get("consideracoes_gerais") or "").strip()
    if cons:
        _analysis_cards(doc, cons)
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

        # Embargos à Execução e/ou Exceção de Pré-Executividade
        _sub_gray(doc, "Embargos e/ou Exceção")
        embargos = cred.get("embargos") or []
        if not embargos:
            embargos = [{}]
        _cont_def = {}
        for emb in embargos:
            tipo = (emb.get("tipo") or "").strip() or "Embargos à Execução"
            _cont_def[tipo] = _cont_def.get(tipo, 0) + 1
            _para(doc, f"{tipo} nº {_cont_def[tipo]}", bold=True, size=9.5, color=_TXT, before=3, after=2)
            _kv_label_table(doc, [
                ("Embargante / Excipiente", emb.get("embargante", "")),
                ("Data da Distribuição",  emb.get("data_dist", "")),
                ("Tese",                  emb.get("tese", "")),
                ("Principais Andamentos", emb.get("andamentos_resumo", "")),
                ("Status Atual",          emb.get("status", "")),
            ])
        _note(doc, "⊕ Replicar o quadro acima para cada embargo / exceção adicional identificado.")

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

    teses = dados.get("teses_recuperacao") or {}

    # 2.1 Penhora Direta
    penhora = teses.get("penhora_direta") or {}
    _sub_orange(doc, "2.1 Penhora Direta")
    _sub_gray(doc, "a. Ponderações e Observações Gerais")
    _analysis_cards(
        doc,
        penhora.get("analise"),
        "Descrever aqui: situação geral dos imóveis localizados, gravames identificados, "
        "bloqueios de matrícula, imóveis não localizados, pesquisas pendentes e demais "
        "observações relevantes.",
    )
    _sub_gray(doc, "b. Matrículas Mapeadas")
    _asset_cards(doc, ativos, tese="Penhora Direta", placeholder=True)
    _spacer(doc)

    # 2.2 IDPJ
    idpj = teses.get("idpj") or {}
    _sub_orange(doc, "2.2 IDPJ")
    _sub_gray(doc, "a. Resumo da Tese")
    _analysis_cards(
        doc,
        idpj.get("resumo"),
        "Fundamento principal · Histórico dos fatos · Objetivo do IDPJ · Principais elementos "
        "probatórios identificados (procurações, representação conjunta, declarações em processo, "
        "confusão patrimonial, etc.).",
    )
    _sub_gray(doc, "b. Empresa-Alvo")
    empresa = idpj.get("empresa_alvo") or {}
    _kv_table(doc, ["CAMPO", "DADO"], [
        ("Razão social da empresa-alvo", empresa.get("razao_social", "")),
        ("CNPJ", empresa.get("cnpj", "")),
        ("CNAE principal", empresa.get("cnae_principal", "")),
        ("Atividades secundárias", empresa.get("atividades_secundarias", "")),
        ("Sócio atual", empresa.get("socio_atual", "")),
        ("Endereço fiscal", empresa.get("endereco_fiscal", "")),
        ("Fundamentação da tese", empresa.get("fundamentacao", "")),
        ("Crédito mais adequado para propositura", empresa.get("credito_propositura", "")),
    ])
    _spacer(doc, pts=2)
    _sub_gray(doc, "c. Cronologia Societária — Atos Relevantes")
    cronologia = idpj.get("cronologia") or []
    _grid_table(
        doc,
        ["DATA", "ATO", "DETALHAMENTO / FONTE"],
        [[
            item.get("data", ""),
            item.get("ato", ""),
            " · ".join(x for x in [item.get("detalhamento", ""), item.get("referencia", "")] if x),
        ] for item in cronologia],
        [2.6, 4.5, 9.8],
        min_rows=2,
    )
    _spacer(doc, pts=2)
    _sub_gray(doc, "d. Evidências de Controle Informal")
    _analysis_cards(
        doc,
        idpj.get("evidencias"),
        "Espaço destinado à consolidação dos principais elementos probatórios que sustentam a "
        "tese: fatos relevantes identificados durante a investigação, documentos, prints, imagens, "
        "pesquisas patrimoniais, alterações societárias, procurações, manifestações processuais e "
        "demais evidências que demonstrem a viabilidade da medida.",
    )
    _sub_gray(doc, "e. Ativos Atingíveis via IDPJ  —  ~R$ [___] (VM) / ~R$ [___] (VP)")
    _asset_cards(doc, ativos, tese="IDPJ", placeholder=True)
    _spacer(doc, pts=2)
    _sub_gray(doc, "f. Upside Identificado (se houver)")
    _analysis_cards(
        doc,
        idpj.get("upside"),
        "Descrever eventual upside — ex: recebíveis futuros de compra e venda em andamento, "
        "participações societárias, créditos a receber, outros ativos. Indicar valor estimado e "
        "prazo de realização.",
    )
    _spacer(doc)

    # 2.3 Fraude à Execução
    fraude = teses.get("fraude_execucao") or {}
    _sub_orange(doc, "2.3 Fraude à Execução")
    _sub_gray(doc, "a. Resumo da Tese")
    _analysis_cards(
        doc,
        fraude.get("resumo"),
        "Descrever os negócios jurídicos impugnáveis (doações, transferências, usufruto) "
        "praticados no curso ou após o ajuizamento da execução. Indicar a data do ajuizamento e a "
        "data de cada ato para enquadramento da tese. Identificar qual crédito é mais adequado "
        "para a propositura.",
    )
    _sub_gray(doc, "b. Má-fé e Insolvência")
    _analysis_cards(
        doc,
        fraude.get("ma_fe_insolvencia"),
        "Registrar os elementos que demonstram a má-fé do adquirente (consilium fraudis) e a "
        "insolvência do devedor decorrente do ato — pressupostos da fraude à execução / ação pauliana.",
    )
    _sub_gray(doc, "c. Ativos Atingíveis via Fraude à Execução  —  ~R$ [___] (VM) / ~R$ [___] (VP)")
    _asset_cards(doc, ativos, tese="Fraude à Execução", placeholder=True)
    _spacer(doc)

    # 2.4 Outras teses
    _sub_orange(doc, "2.4 Outras teses")
    _analysis_cards(
        doc,
        teses.get("outras"),
        "Registrar outras teses de recuperação eventualmente aplicáveis ao caso.",
    )
    _spacer(doc)

    # ══ 3. PASSIVO ═══════════════════════════════════════════════════════
    _sec_title(doc, "3. PASSIVO")
    _note(doc, "Mapear o passivo integral dos devedores e das empresas-alvo. O passivo é fator "
               "determinante na avaliação do risco real de recuperação.")
    passivo = dados.get("passivo") or {}

    # 3.1 Passivo Fiscal
    _sub_orange(doc, "3.1 Passivo Fiscal")
    _sub_gray(doc, "e-CAC")
    e_cac = passivo.get("e_cac") or []
    _grid_table(
        doc,
        ["NOME", "CPF / CNPJ", "SALDO e-CAC (R$)"],
        [[item.get("nome", ""), item.get("cpf_cnpj", ""), item.get("saldo", "")] for item in e_cac],
        [7.0, 5.0, 4.9],
        min_rows=2,
    )
    _spacer(doc, pts=2)
    _sub_gray(doc, "Execuções Fiscais")
    _process_cards(doc, passivo.get("fiscais") or [], "Exequente")
    _spacer(doc)

    # 3.2 Passivo Trabalhista
    _sub_orange(doc, "3.2 Passivo Trabalhista")
    _process_cards(doc, passivo.get("trabalhistas") or [], "Autor / Reclamante")
    _spacer(doc)

    # 3.3 Passivo Cível (Terceiros)
    _sub_orange(doc, "3.3 Passivo Cível (Terceiros)")
    _process_cards(doc, passivo.get("civeis") or [], "Parte contrária")
    _spacer(doc, pts=2)

    # Principais Credores
    _sub_gray(doc, "Principais Credores")
    credores = passivo.get("principais_credores") or []
    if credores:
        for i, credor in enumerate(credores):
            _kv_table(doc, (credor.get("credor") or f"Credor nº {i + 1}").upper(), [
                ("Natureza do crédito", credor.get("natureza", "")),
                ("Valor estimado", credor.get("valor", "")),
                ("Garantia / preferência", credor.get("garantia_preferencia", "")),
                ("Status", credor.get("status", "")),
                ("Observações", credor.get("observacoes", "")),
            ], w_label=4.3, w_value=12.6)
            if i < len(credores) - 1:
                _spacer(doc, pts=3)
    else:
        _kv_table(doc, ["CREDOR", "INFORMAÇÃO"], [
            ("Natureza do crédito", ""),
            ("Valor estimado", ""),
            ("Garantia / preferência", ""),
            ("Status e observações", ""),
        ], w_label=4.3, w_value=12.6)
    _sub_gray(doc, "Pontos de Atenção sobre o Passivo")
    _analysis_cards(
        doc,
        passivo.get("pontos_atencao"),
        "Hipotecas preferentes · créditos arrematados por terceiros · passivo oculto potencial · "
        "ações de reintegração · litígios relevantes.",
    )
    _spacer(doc)

    # ══ 5. PENDÊNCIAS ════════════════════════════════════════════════════
    _sec_title(doc, "5. PENDÊNCIAS")
    _note(doc, "Campo destinado ao registro de pendências relacionadas à pesquisa de bens, obtenção de "
               "escrituras, avaliações de imóveis e demais diligências ainda necessárias para a conclusão "
               "da análise do caso.")
    pendencias = dados.get("pendencias") or []
    _grid_table(
        doc,
        ["PENDÊNCIA", "PROTOCOLO", "PRAZO / STATUS"],
        [[p.get("pendencia", ""), p.get("protocolo", ""), p.get("prazo_status", "")] for p in pendencias],
        [8.9, 4.0, 4.0],
        min_rows=2,
    )
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

def gerar_dossie_word(texto_completo: str, relatorio: str, client, model: str) -> str:
    """Extrai dados do texto OCR completo (+ relatório de apoio) e gera o Word PPA v3.0."""
    dados = _extrair_dados(texto_completo, relatorio, client, model)
    # Passada final: confere campos vazios contra o relatório e completa o que faltou
    dados = _completar_com_relatorio(dados, relatorio, client, model)
    return _build_doc(dados)


# ══════════════════════════════════════════════════════════════════════════
# Preenchimento do PASSIVO (Seção 3) a partir de dados da Predictus
# ══════════════════════════════════════════════════════════════════════════

def _fmt_valor_br(v) -> str:
    if v is None or v == "":
        return ""
    try:
        return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def _proc_to_row(p: dict) -> list:
    """Mapeia um processo Predictus para as 7 colunas do passivo do dossiê."""
    return [
        p.get("cnj", ""),                 # Nº CNJ
        p.get("vinc", ""),                # Vinculado a
        p.get("data", "") or "",          # Distribuição
        p.get("ativo", ""),               # Exequente / Autor (polo ativo)
        _fmt_valor_br(p.get("valor")),    # Valor Causa (R$)
        "",                               # SAT Est. (R$) — não consta na Predictus
        p.get("status", ""),              # Status / Obs.
    ]


def _fill_passivo_table(table, registros: list, parte_label: str):
    """Preenche o passivo novo em cartões e mantém compatibilidade com o modelo antigo."""
    ncols = len(table.columns)
    if ncols == 2:
        _fill_process_cards(table, registros, parte_label)
        return
    larguras = [c.width for c in table.rows[0].cells]
    # remove linhas de dados (mantém a primeira = header laranja)
    for row in list(table.rows)[1:]:
        table._tbl.remove(row._tr)
    if not registros:
        # deixa uma linha vazia para não colar o próximo bloco no header
        registros = [[""] * ncols]
    for reg in registros:
        r = table.add_row()
        for i in range(ncols):
            c = r.cells[i]
            _set_bg(c, _BRANCO); _set_borders(c); _cell_pad(c)
            _write(c, reg[i] if i < len(reg) else "")
            if i < len(larguras) and larguras[i]:
                c.width = larguras[i]


def _iter_headings_tables(doc):
    """Itera o corpo em ordem, associando cada tabela ao último subtítulo não-vazio."""
    last_heading = ""
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            txt = _DocxParagraph(child, doc).text.strip()
            if txt:
                last_heading = txt
        elif child.tag == qn("w:tbl"):
            yield last_heading, _DocxTable(child, doc)


def preencher_passivo_dossie(dossie_path, fiscais: list, trabalhistas: list, civeis: list) -> str:
    """
    Preenche as tabelas da Seção 3 (Passivo) do dossiê — Execuções Fiscais,
    Trabalhista e Cível — com os processos consolidados da Predictus.
    Não toca em 'Principais Credores' nem no e-CAC.
    Se dossie_path for None, gera um dossiê novo (esqueleto) e preenche nele.
    """
    if dossie_path:
        doc = Document(dossie_path)
    else:
        doc = Document(_build_doc({}))

    for heading, table in _iter_headings_tables(doc):
        h = heading.lower()
        if "execuções fiscais" in h or "execucoes fiscais" in h:
            _fill_passivo_table(table, [_proc_to_row(p) for p in fiscais], "Exequente")
        elif "trabalhista" in h:
            _fill_passivo_table(table, [_proc_to_row(p) for p in trabalhistas], "Autor / Reclamante")
        elif "cível" in h or "civel" in h:
            _fill_passivo_table(table, [_proc_to_row(p) for p in civeis], "Parte contrária")

    caminho = os.path.join(tempfile.gettempdir(), "Dossie_PPA_atualizado.docx")
    doc.save(caminho)
    return caminho
