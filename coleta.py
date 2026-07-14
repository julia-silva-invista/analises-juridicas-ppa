# -*- coding: utf-8 -*-
"""
Coleta de Informações — preenche planilha x.xlsx com dados da Predictus
"""

import os
import re
import shutil
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from typing import Optional, Tuple, Any

from dossie_ppa import preencher_passivo_dossie

# ── Constantes ───────────────────────────────────────────────────────────────

TEMPLATE_PATH = "x.xlsx"

CLASSE_MAP = {
    "EXECUCAO DE TITULO EXTRAJUDICIAL": "Execução de Título Extrajudicial",
    "EMBARGOS A EXECUCAO": "Embargos à Execução",
    "EMBARGOS À EXECUÇÃO": "Embargos à Execução",
    "CARTA PRECATORIA CIVEL": "Carta Precatória Cível",
    "CARTA PRECATÓRIA CÍVEL": "Carta Precatória Cível",
    "AGRAVO DE INSTRUMENTO": "Agravo de Instrumento",
    "ACAO TRABALHISTA - RITO ORDINARIO": "Ação Trabalhista — Rito Ordinário",
    "AÇÃO TRABALHISTA - RITO ORDINÁRIO": "Ação Trabalhista — Rito Ordinário",
    "ACAO TRABALHISTA - RITO SUMARIÍSSIMO": "Ação Trabalhista — Rito Sumaríssimo",
    "ACAO TRABALHISTA - RITO SUMARÍSSIMO": "Ação Trabalhista — Rito Sumaríssimo",
    "ACAO CIVIL PUBLICA": "Ação Civil Pública",
    "AÇÃO CIVIL PÚBLICA": "Ação Civil Pública",
    "EXECUCAO FISCAL": "Execução Fiscal",
    "MANDADO DE SEGURANCA": "Mandado de Segurança",
    "MANDADO DE SEGURANÇA": "Mandado de Segurança",
    "ACAO DECLARATORIA": "Ação Declaratória",
    "AÇÃO DECLARATÓRIA": "Ação Declaratória",
    "MONITORIA": "Monitória",
    "MONITÓRIA": "Monitória",
    "CUMPRIMENTO DE SENTENCA": "Cumprimento de Sentença",
    "CUMPRIMENTO DE SENTENÇA": "Cumprimento de Sentença",
    "ACAO DE COBRANCA": "Ação de Cobrança",
    "AÇÃO DE COBRANÇA": "Ação de Cobrança",
    "APELACAO CIVEL": "Apelação Cível",
    "APELAÇÃO CÍVEL": "Apelação Cível",
    "EMBARGOS DE DECLARACAO": "Embargos de Declaração",
    "EMBARGOS DE DECLARAÇÃO": "Embargos de Declaração",
    "RECURSO DE REVISTA": "Recurso de Revista",
    "RECURSO ORDINARIO TRABALHISTA": "Recurso Ordinário Trabalhista",
    "RECURSO ORDINÁRIO TRABALHISTA": "Recurso Ordinário Trabalhista",
    "ACAO RESCISORIA": "Ação Rescisória",
    "AÇÃO RESCISÓRIA": "Ação Rescisória",
    "RECLAMACAO TRABALHISTA": "Reclamação Trabalhista",
    "RECLAMAÇÃO TRABALHISTA": "Reclamação Trabalhista",
    "DISSIDIO INDIVIDUAL TRABALHISTA": "Dissídio Individual Trabalhista",
    "DISSÍDIO INDIVIDUAL TRABALHISTA": "Dissídio Individual Trabalhista",
    "ACAO DE INDENIZACAO POR DANO MORAL": "Ação de Indenização por Dano Moral",
    "AÇÃO DE INDENIZAÇÃO POR DANO MORAL": "Ação de Indenização por Dano Moral",
}

STATUS_MAP = {
    "EM TRAMITACAO": "Em tramitação",
    "EM TRAMITAÇÃO": "Em tramitação",
    "ARQUIVADO": "Arquivado definitivamente",
    "ARQUIVADO DEFINITIVAMENTE": "Arquivado definitivamente",
    "BAIXADO": "Arquivado definitivamente",
    "SUSPENSO": "Suspenso",
    "EXTINTO": "Extinto",
    "JULGADO": "Julgado",
    "TRANSITADO EM JULGADO": "Transitado em julgado",
}

_LOWERCASE_WORDS = {
    "de", "da", "do", "das", "dos", "e", "em", "a", "o", "ao", "à",
    "na", "no", "nas", "nos", "por", "para", "com", "sem", "sob",
    "sobre", "entre",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def _title_case(text: str) -> str:
    if not text:
        return ""
    words = str(text).strip().split()
    result = []
    for i, word in enumerate(words):
        if i == 0 or word.lower() not in _LOWERCASE_WORDS:
            result.append(word.capitalize())
        else:
            result.append(word.lower())
    return " ".join(result)


def _format_document(raw: str) -> str:
    d = re.sub(r"\D", "", raw)
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return raw


def _doc_label(raw: str) -> str:
    d = re.sub(r"\D", "", raw)
    if len(d) == 11:
        return "CPF"
    if len(d) == 14:
        return "CNPJ"
    return ""


def _format_parties(party_str: str) -> str:
    raw = str(party_str).strip()
    if raw in ("", "nan", "None"):
        return ""
    parties = []
    for part in raw.split("|"):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            part = part.split(":", 1)[1].strip()
        m = re.search(r"\[(\d+)\]", part)
        name = re.sub(r"\[.*?\]", "", part).strip()
        name = _title_case(name)
        if not name:
            continue
        if m:
            doc_raw = m.group(1)
            label = _doc_label(doc_raw)
            doc_fmt = _format_document(doc_raw)
            if label:
                parties.append(f"{name} ({label} nº {doc_fmt})")
            else:
                parties.append(name)
        else:
            parties.append(name)
    return "; ".join(parties)


def _format_classe(classe: str) -> str:
    s = str(classe).strip()
    if s in ("", "nan", "None"):
        return ""
    return CLASSE_MAP.get(s.upper(), _title_case(s))


def _format_status(status: str) -> str:
    s = str(status).strip()
    if s in ("", "nan", "None"):
        return ""
    return STATUS_MAP.get(s.upper(), _title_case(s))


def _parse_date(val) -> Optional[str]:
    s = str(val).strip()
    if s in ("", "nan", "None", "NaT"):
        return None
    if re.match(r"\d{2}/\d{2}/\d{4}", s):
        return s
    try:
        from dateutil import parser as dp
        return dp.parse(s, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return s


def _is_trabalhista(ramo: str) -> bool:
    r = str(ramo).upper()
    return any(kw in r for kw in ("TRABALHO", "TRABALHISTA"))


def _is_fiscal(ramo: str, classe: str) -> bool:
    s = (str(ramo) + " " + str(classe)).upper()
    return any(kw in s for kw in (
        "FISCAL", "TRIBUT", "DIVIDA ATIVA", "DÍVIDA ATIVA", "EXECUCAO FISCAL", "EXECUÇÃO FISCAL",
    ))


def _get_nome_from_filename(filepath: str) -> str:
    base = os.path.splitext(os.path.basename(filepath))[0]
    return re.sub(r"^Predictus\s*[-–—]\s*", "", base, flags=re.IGNORECASE).strip()


def _first_empty_row(ws, col: int = 1) -> int:
    for r in range(2, ws.max_row + 2):
        if ws.cell(row=r, column=col).value is None:
            return r
    return ws.max_row + 1


def _col(df: pd.DataFrame, keywords: list) -> str:
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in col.lower():
                return col
    return ""


# ── Função principal ──────────────────────────────────────────────────────────

def coleta_gerar(excel_files) -> Tuple[str, str, Any]:
    """
    Lê Excel(s) da Predictus e preenche o template x.xlsx.
    Retorna: (log_text, status_msg, caminho_do_arquivo_ou_None)
    """
    lines: list = []

    def log(msg: str):
        lines.append(msg)
        return "\n".join(lines)

    if not excel_files:
        return "Nenhum arquivo enviado.", "Erro: nenhum arquivo", None

    if not os.path.exists(TEMPLATE_PATH):
        return (
            f"Template '{TEMPLATE_PATH}' não encontrado.\n"
            "Certifique-se de que o arquivo x.xlsx está na mesma pasta que os arquivos .py.",
            "Erro: template ausente",
            None,
        )

    os.makedirs("resultados", exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = f"resultados/coleta_{ts}.xlsx"
    shutil.copy2(TEMPLATE_PATH, out_path)

    wb = load_workbook(out_path)
    ws_trab   = wb["Trabalhista"]
    ws_fiscal = wb["Fiscal & Cível"]
    ws_capa   = wb["Capa"]

    row_trab   = _first_empty_row(ws_trab)
    row_fiscal = _first_empty_row(ws_fiscal)
    row_capa   = _first_empty_row(ws_capa)

    total_trab = total_fiscal = 0

    for file_obj in excel_files:
        fp = file_obj.name if hasattr(file_obj, "name") else str(file_obj)
        log(f"📂 {os.path.basename(fp)}")

        try:
            xl = pd.ExcelFile(fp)
            sheet = "Dossiê Jurídico" if "Dossiê Jurídico" in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(fp, sheet_name=sheet, dtype=str)
        except Exception as exc:
            log(f"   ❌ Erro ao ler arquivo: {exc}")
            continue

        if df.empty:
            log("   ⚠️ Planilha vazia.")
            continue

        df.columns = [c.strip() for c in df.columns]

        nome = _get_nome_from_filename(fp)
        cpf_col = _col(df, ["termo buscado", "termo", "buscado", "cpf", "cnpj"])
        cpf = str(df[cpf_col].iloc[0]) if cpf_col else ""

        log(f"   👤 {nome}  |  {cpf}")
        log(f"   📋 {len(df)} processo(s)")

        ws_capa.cell(row=row_capa, column=1, value=nome)
        ws_capa.cell(row=row_capa, column=3, value=cpf)
        row_capa += 1

        col_ramo    = _col(df, ["ramo do direito", "ramo"])
        col_valor   = _col(df, ["valor da causa", "valor"])
        col_cnj     = _col(df, ["n° processo", "nº processo", "numero", "processo"])
        col_data    = _col(df, ["data de distribui", "distribuição", "distribuicao"])
        col_classe  = _col(df, ["classe processual", "classe"])
        col_ativo   = _col(df, ["partes ativas", "polo ativo"])
        col_passivo = _col(df, ["partes passivas", "polo passivo"])
        col_status  = _col(df, ["status", "situação", "situacao"])

        n_t = n_f = 0

        for _, row in df.iterrows():

            def g(col_name: str) -> str:
                if not col_name:
                    return ""
                v = row.get(col_name, "")
                return str(v) if v and str(v) not in ("nan", "None") else ""

            ramo = g(col_ramo)
            valor_raw = g(col_valor)
            try:
                valor = float(valor_raw.replace(",", ".")) if valor_raw else None
            except Exception:
                valor = None

            proc = {
                "cnj":     g(col_cnj),
                "vinc":    nome,
                "data":    _parse_date(g(col_data)),
                "classe":  _format_classe(g(col_classe)),
                "valor":   valor,
                "ativo":   _format_parties(g(col_ativo)),
                "passivo": _format_parties(g(col_passivo)),
                "status":  _format_status(g(col_status)),
            }

            if _is_trabalhista(ramo):
                r, ws = row_trab, ws_trab
                row_trab += 1
                n_t += 1
            else:
                r, ws = row_fiscal, ws_fiscal
                row_fiscal += 1
                n_f += 1

            ws.cell(row=r, column=1, value=proc["cnj"])
            ws.cell(row=r, column=2, value=proc["vinc"])
            ws.cell(row=r, column=3, value=proc["data"])
            ws.cell(row=r, column=4, value=proc["classe"])
            ws.cell(row=r, column=5, value=proc["valor"])
            # coluna 6 = Saldo Atualizado (fórmula — não sobrescrever)
            ws.cell(row=r, column=7, value=proc["ativo"])
            ws.cell(row=r, column=8, value=proc["passivo"])
            ws.cell(row=r, column=9, value=proc["status"])

        total_trab  += n_t
        total_fiscal += n_f
        log(f"   ✅ {n_t} trabalhista(s) | {n_f} fiscal/cível")

    wb.save(out_path)

    total = total_trab + total_fiscal
    log(f"\n✅ Concluído! {total} processo(s) preenchido(s):")
    log(f"   • Trabalhista:    {total_trab}")
    log(f"   • Fiscal & Cível: {total_fiscal}")

    status_msg = f"{total} processos preenchidos — {total_trab} trabalhistas, {total_fiscal} fiscal/cível"
    return "\n".join(lines), status_msg, out_path


# ── Dossiê atualizado: passivo (Seção 3) a partir da Predictus ────────────────

def _classificar_predictus(excel_files, log) -> Tuple[list, list, list]:
    """Lê os Excel(s) da Predictus e separa os processos em fiscal / trabalhista / cível."""
    fiscais, trabalhistas, civeis = [], [], []
    for file_obj in excel_files:
        fp = file_obj.name if hasattr(file_obj, "name") else str(file_obj)
        log(f"📂 {os.path.basename(fp)}")
        try:
            xl = pd.ExcelFile(fp)
            sheet = "Dossiê Jurídico" if "Dossiê Jurídico" in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(fp, sheet_name=sheet, dtype=str)
        except Exception as exc:
            log(f"   ❌ Erro ao ler arquivo: {exc}")
            continue
        if df.empty:
            log("   ⚠️ Planilha vazia.")
            continue
        df.columns = [c.strip() for c in df.columns]
        nome = _get_nome_from_filename(fp)

        col_ramo    = _col(df, ["ramo do direito", "ramo"])
        col_valor   = _col(df, ["valor da causa", "valor"])
        col_cnj     = _col(df, ["n° processo", "nº processo", "numero", "processo"])
        col_data    = _col(df, ["data de distribui", "distribuição", "distribuicao"])
        col_classe  = _col(df, ["classe processual", "classe"])
        col_ativo   = _col(df, ["partes ativas", "polo ativo"])
        col_status  = _col(df, ["status", "situação", "situacao"])

        n_f = n_t = n_c = 0
        for _, row in df.iterrows():
            def g(col_name: str) -> str:
                if not col_name:
                    return ""
                v = row.get(col_name, "")
                return str(v) if v and str(v) not in ("nan", "None") else ""

            ramo = g(col_ramo)
            classe = g(col_classe)
            valor_raw = g(col_valor)
            try:
                valor = float(valor_raw.replace(",", ".")) if valor_raw else None
            except Exception:
                valor = None

            proc = {
                "cnj":    g(col_cnj),
                "vinc":   nome,
                "data":   _parse_date(g(col_data)) or "",
                "ativo":  _format_parties(g(col_ativo)),
                "valor":  valor,
                "status": _format_status(g(col_status)),
            }
            if _is_trabalhista(ramo):
                trabalhistas.append(proc); n_t += 1
            elif _is_fiscal(ramo, classe):
                fiscais.append(proc); n_f += 1
            else:
                civeis.append(proc); n_c += 1
        log(f"   👤 {nome} — {n_f} fiscal | {n_t} trabalhista | {n_c} cível")
    return fiscais, trabalhistas, civeis


def coleta_gerar_dossie(excel_files, dossie_file):
    """
    Preenche a Seção 3 (Passivo) de um dossiê PPA com os processos da Predictus.
    Retorna: (log_text, status_msg, caminho_do_docx_ou_None).
    """
    lines: list = []

    def log(msg: str):
        lines.append(msg)
        return "\n".join(lines)

    if not excel_files:
        return "Nenhum Excel da Predictus enviado.", "Erro: nenhum Excel da Predictus", None

    try:
        fiscais, trabalhistas, civeis = _classificar_predictus(excel_files, log)
    except Exception as exc:
        return log(f"\n❌ Erro ao ler a Predictus: {exc}"), "Erro na leitura da Predictus", None

    total = len(fiscais) + len(trabalhistas) + len(civeis)
    if total == 0:
        return log("\n⚠️ Nenhum processo identificado nos arquivos."), "Nenhum processo identificado", None

    dossie_path = None
    if dossie_file:
        dossie_path = dossie_file.name if hasattr(dossie_file, "name") else str(dossie_file)
        log(f"\n📄 Atualizando dossiê enviado: {os.path.basename(dossie_path)}")
    else:
        log("\n📄 Nenhum dossiê enviado — gerando esqueleto com o passivo preenchido.")

    try:
        out = preencher_passivo_dossie(dossie_path, fiscais, trabalhistas, civeis)
    except Exception as exc:
        return log(f"\n❌ Erro ao preencher o dossiê: {exc}"), "Erro ao preencher o dossiê", None

    log(f"\n✅ Passivo preenchido: {len(fiscais)} fiscal · {len(trabalhistas)} trabalhista · {len(civeis)} cível.")
    status = f"Passivo preenchido — {len(fiscais)} fiscal, {len(trabalhistas)} trabalhista, {len(civeis)} cível"
    return "\n".join(lines), status, out
