# -*- coding: utf-8 -*-
"""Extração, edição e exportação da Timeline Societária."""

from __future__ import annotations

import copy
import html
import json
import os
import re
import tempfile
import time
import unicodedata
from pathlib import Path
from typing import Any

import gradio as gr
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
from PIL import Image, ImageDraw, ImageFont

from utils import _retry


MODEL_TIMELINE = os.getenv("GEMINI_MODEL_TIMELINE", "gemini-2.5-pro")
TEMPLATE_PATH = Path(__file__).parent / "templates" / "cronologia_societaria.docx"
EDITOR_HEADERS = [
    "Data",
    "Ato / ACS",
    "Detalhamento",
    "Sócios após o ato (Nome | %)",
    "Administração",
    "Capital social",
    "Sede",
    "Objeto social",
    "Cessões de quotas (Cedente > Cessionário | % | valor)",
    "Imóveis (Matrícula | Cartório | Valor | Movimento)",
    "Filiais",
    "Fonte",
]

PROMPT_TIMELINE = """\
Você é um analista jurídico societário sênior. Analise TODOS os atos societários
anexados, organize-os cronologicamente e reconstrua o estado da sociedade após
cada ato. Aplique OCR visual quando o PDF estiver escaneado.

REGRAS CRÍTICAS
- Não invente. Quando o documento não informar algo, use string vazia ou lista vazia.
- Diferencie data de assinatura, data do ato e data de arquivamento. Em "data",
  priorize a data do ato; se ausente, use a data do arquivamento.
- "ato" deve trazer a denominação legível: Constituição, 1ª Alteração,
  5ª Alteração/ACS etc. "numero_arquivamento" é o número da Junta Comercial.
- Em "socios_apos", represente o quadro societário resultante DEPOIS do ato.
  Informe percentual quando constar ou puder ser calculado exatamente pelas quotas.
- Em "administradores_apos", represente a administração resultante depois do ato.
- Em "capital_social_apos" e "capital_social_anterior", use o valor em reais com
  formatação brasileira. Preencha o anterior sempre que o documento informar.
- Em "sede_apos", informe cidade/UF da sede depois do ato. Em "objeto_apos",
  resuma o objeto social vigente depois do ato (uma linha).
- Em "cessoes", registre CADA transmissão de quotas: quem cedeu ("cedente"),
  quem recebeu ("cessionario"), o percentual ou número de quotas cedidas
  ("participacao"), o valor da cessão ("valor") e observações relevantes
  ("observacao") — por exemplo valor simbólico, cessão a parente ou a pessoa
  jurídica do mesmo grupo, doação, permuta.
- Em "imoveis", registre CADA bem imóvel movimentado no ato, com "matricula"
  (número da matrícula), "cartorio" (registro de imóveis), "cidade", "valor" e
  "movimento" preenchido com uma destas palavras: "integralizacao" (conferido ao
  capital), "saida" (restituído ao sócio, cindido ou transferido), "aquisicao"
  (comprado pela sociedade). Em "descricao", detalhe a operação.
- Em "filiais_apos", liste todas as filiais existentes depois do ato; em
  "filiais_adicionadas", somente as criadas naquele ato.
- Em "detalhamento", descreva objetivamente todas as mudanças relevantes:
  entrada/saída de sócio, cessão de quotas, integralização, capital, administração,
  objeto, sede, nome empresarial, filiais, dissolução ou outras deliberações.
- Cite a fonte em linguagem humana: nome do documento e página(s).
- Ordene os eventos em ordem cronológica crescente.

Responda SOMENTE com JSON válido nesta estrutura:
{
  "empresa": "",
  "cnpj": "",
  "eventos": [
    {
      "data": "DD/MM/AAAA",
      "ato": "",
      "numero_arquivamento": "",
      "detalhamento": "",
      "categorias": ["Sócios", "Capital", "Administração", "Filial", "Imóvel", "Sede", "Objeto", "Outros"],
      "socios_apos": [{"nome": "", "participacao": "", "quotas": ""}],
      "administradores_apos": [{"nome": "", "cargo": "Administrador"}],
      "capital_social_anterior": "",
      "capital_social_apos": "",
      "sede_apos": "",
      "objeto_apos": "",
      "cessoes": [{"cedente": "", "cessionario": "", "participacao": "", "valor": "", "observacao": ""}],
      "imoveis": [{"matricula": "", "cartorio": "", "cidade": "", "valor": "", "movimento": "", "descricao": ""}],
      "filiais_apos": [{"nome": "", "local": ""}],
      "filiais_adicionadas": [{"nome": "", "local": ""}],
      "fonte": ""
    }
  ]
}
"""


def _client() -> genai.Client:
    key = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
    if not key:
        raise RuntimeError("GEMINI_API_KEY_1 não configurada.")
    return genai.Client(
        api_key=key,
        http_options=types.HttpOptions(timeout=int(os.getenv("GEMINI_TIMEOUT_MS", "600000"))),
    )


def _file_path(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        return value.get("path") or value.get("name") or ""
    return getattr(value, "path", None) or getattr(value, "name", None) or ""


def _clean_json(raw: str) -> dict:
    raw = re.sub(r"^```[a-z]*\s*", "", (raw or "").strip(), flags=re.IGNORECASE)
    raw = re.sub(r"\s*```$", "", raw)
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("A resposta da IA não contém um objeto JSON.")
    data.setdefault("empresa", "")
    data.setdefault("cnpj", "")
    data.setdefault("eventos", [])
    data["eventos"] = sorted(
        [e for e in data["eventos"] if isinstance(e, dict)],
        key=lambda e: _date_key(e.get("data", "")),
    )
    return data


def _date_key(value: str) -> tuple[int, int, int]:
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(value))
    if not m:
        return (9999, 12, 31)
    return (int(m.group(3)), int(m.group(2)), int(m.group(1)))


def _upload_pdf(client: genai.Client, path: str):
    upload_path = path
    tmp_ascii = None
    try:
        path.encode("ascii")
    except UnicodeEncodeError:
        import shutil

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.close()
        shutil.copy2(path, tmp.name)
        upload_path = tmp_ascii = tmp.name

    uploaded = client.files.upload(file=upload_path)
    waited = 0
    while waited < 120:
        state = getattr(getattr(uploaded, "state", None), "name", None)
        if state in ("ACTIVE", "FAILED"):
            break
        time.sleep(2)
        waited += 2
        uploaded = client.files.get(name=uploaded.name)
    if tmp_ascii:
        try:
            os.remove(tmp_ascii)
        except OSError:
            pass
    if getattr(getattr(uploaded, "state", None), "name", None) == "FAILED":
        raise RuntimeError(f"Falha ao enviar {Path(path).name} para análise.")
    return uploaded


def _extract(files: list[Any]) -> dict:
    paths = [_file_path(f) for f in (files or [])]
    paths = [p for p in paths if p and Path(p).exists()]
    if not paths:
        raise ValueError("Envie ao menos um ato societário em PDF.")

    client = _client()
    uploaded = []
    try:
        parts: list[types.Part] = [types.Part(text=PROMPT_TIMELINE)]
        for index, path in enumerate(paths, 1):
            item = _upload_pdf(client, path)
            uploaded.append(item)
            parts.append(types.Part(text=f"\nDOCUMENTO {index}: {Path(path).name}\n"))
            parts.append(
                types.Part(
                    file_data=types.FileData(
                        file_uri=item.uri,
                        mime_type=getattr(item, "mime_type", None) or "application/pdf",
                    )
                )
            )

        config = types.GenerateContentConfig(response_mime_type="application/json")

        def _call():
            return client.models.generate_content(
                model=MODEL_TIMELINE,
                contents=[types.Content(role="user", parts=parts)],
                config=config,
            ).text

        return _clean_json(_retry(_call, tentativas=3, espera_base=8))
    finally:
        for item in uploaded:
            try:
                client.files.delete(name=item.name)
            except Exception:
                pass


def _joined_people(items: list[dict], include_share: bool = False) -> str:
    values = []
    for item in items or []:
        name = str(item.get("nome", "")).strip()
        if not name:
            continue
        if include_share:
            share = str(item.get("participacao", "") or item.get("quotas", "")).strip()
            values.append(f"{name} | {share}" if share else name)
        else:
            role = str(item.get("cargo", "")).strip()
            values.append(f"{name} ({role})" if role else name)
    return "; ".join(values)


def _joined_filiais(items: list[dict]) -> str:
    return "; ".join(
        " — ".join(x for x in [str(i.get("nome", "")).strip(), str(i.get("local", "")).strip()] if x)
        for i in (items or [])
        if i.get("nome") or i.get("local")
    )


def _joined_cessoes(items: list[dict]) -> str:
    values = []
    for item in items or []:
        cedente = str(item.get("cedente", "")).strip()
        cessionario = str(item.get("cessionario", "")).strip()
        if not (cedente or cessionario):
            continue
        extra = " | ".join(
            x
            for x in [
                str(item.get("participacao", "")).strip(),
                str(item.get("valor", "")).strip(),
                str(item.get("observacao", "")).strip(),
            ]
            if x
        )
        base = f"{cedente} > {cessionario}"
        values.append(f"{base} | {extra}" if extra else base)
    return "; ".join(values)


def _parse_cessoes(value: str) -> list[dict]:
    result = []
    for chunk in re.split(r"\s*;\s*", str(value or "").strip()):
        if not chunk:
            continue
        parts = [p.strip() for p in chunk.split("|")]
        pessoas = re.split(r"\s*(?:>|→|para)\s*", parts[0], maxsplit=1)
        result.append(
            {
                "cedente": pessoas[0].strip(),
                "cessionario": pessoas[1].strip() if len(pessoas) > 1 else "",
                "participacao": parts[1] if len(parts) > 1 else "",
                "valor": parts[2] if len(parts) > 2 else "",
                "observacao": " | ".join(parts[3:]) if len(parts) > 3 else "",
            }
        )
    return result


def _joined_imoveis(items: list[dict]) -> str:
    values = []
    for item in items or []:
        matricula = str(item.get("matricula", "")).strip()
        if not matricula and not item.get("cartorio"):
            continue
        values.append(
            " | ".join(
                x
                for x in [
                    matricula,
                    str(item.get("cartorio", "")).strip() or str(item.get("cidade", "")).strip(),
                    str(item.get("valor", "")).strip(),
                    str(item.get("movimento", "")).strip(),
                ]
                if x
            )
        )
    return "; ".join(values)


def _parse_imoveis(value: str) -> list[dict]:
    result = []
    for chunk in re.split(r"\s*;\s*", str(value or "").strip()):
        if not chunk:
            continue
        parts = [p.strip() for p in chunk.split("|")]
        parts += [""] * (4 - len(parts))
        result.append(
            {
                "matricula": parts[0],
                "cartorio": parts[1],
                "cidade": parts[1],
                "valor": parts[2],
                "movimento": parts[3] or "integralizacao",
                "descricao": "",
            }
        )
    return result


def data_to_rows(data: dict) -> list[list[str]]:
    rows = []
    for event in data.get("eventos", []):
        ato = str(event.get("ato", "")).strip()
        number = str(event.get("numero_arquivamento", "")).strip()
        if number and number not in ato:
            ato = f"{ato} | Arquiv. {number}" if ato else f"Arquiv. {number}"
        rows.append(
            [
                str(event.get("data", "")),
                ato,
                str(event.get("detalhamento", "")),
                _joined_people(event.get("socios_apos", []), include_share=True),
                _joined_people(event.get("administradores_apos", [])),
                str(event.get("capital_social_apos", "")),
                str(event.get("sede_apos", "")),
                str(event.get("objeto_apos", "")),
                _joined_cessoes(event.get("cessoes", [])),
                _joined_imoveis(event.get("imoveis", [])),
                _joined_filiais(event.get("filiais_apos", [])),
                str(event.get("fonte", "")),
            ]
        )
    return rows


def _parse_people(value: str, administrators: bool = False) -> list[dict]:
    result = []
    for chunk in re.split(r"\s*;\s*", str(value or "").strip()):
        if not chunk:
            continue
        if administrators:
            m = re.match(r"(.+?)\s*\(([^()]*)\)\s*$", chunk)
            result.append(
                {
                    "nome": (m.group(1) if m else chunk).strip(),
                    "cargo": (m.group(2) if m else "Administrador").strip(),
                }
            )
        else:
            parts = [p.strip() for p in chunk.split("|", 1)]
            result.append(
                {
                    "nome": parts[0],
                    "participacao": parts[1] if len(parts) > 1 else "",
                    "quotas": "",
                }
            )
    return result


def _parse_filiais(value: str) -> list[dict]:
    result = []
    for chunk in re.split(r"\s*;\s*", str(value or "").strip()):
        if not chunk:
            continue
        parts = re.split(r"\s+[—-]\s+", chunk, maxsplit=1)
        result.append({"nome": parts[0].strip(), "local": parts[1].strip() if len(parts) > 1 else ""})
    return result


def rows_to_data(rows: Any, previous: dict | None = None) -> dict:
    if hasattr(rows, "values"):
        rows = rows.values.tolist()
    anteriores = {
        str(e.get("data", "")) + "|" + str(e.get("ato", "")): e
        for e in (previous or {}).get("eventos", [])
    }
    result = {
        "empresa": (previous or {}).get("empresa", ""),
        "cnpj": (previous or {}).get("cnpj", ""),
        "eventos": [],
    }
    for row in rows or []:
        values = list(row) + [""] * (len(EDITOR_HEADERS) - len(row))
        if not any(str(v or "").strip() for v in values):
            continue
        ato_raw = str(values[1] or "").strip()
        number = ""
        m = re.search(r"\|\s*Arquiv\.\s*(.+)$", ato_raw, flags=re.IGNORECASE)
        if m:
            number = m.group(1).strip()
            ato_raw = ato_raw[: m.start()].strip()
        anterior = anteriores.get(str(values[0] or "").strip() + "|" + ato_raw, {})
        result["eventos"].append(
            {
                "data": str(values[0] or "").strip(),
                "ato": ato_raw,
                "numero_arquivamento": number,
                "detalhamento": str(values[2] or "").strip(),
                "categorias": [],
                "socios_apos": _parse_people(values[3]),
                "administradores_apos": _parse_people(values[4], administrators=True),
                "capital_social_anterior": str(anterior.get("capital_social_anterior", "")).strip(),
                "capital_social_apos": str(values[5] or "").strip(),
                "sede_apos": str(values[6] or "").strip(),
                "objeto_apos": str(values[7] or "").strip(),
                "cessoes": _parse_cessoes(values[8]),
                "imoveis": _parse_imoveis(values[9]),
                "filiais_apos": _parse_filiais(values[10]),
                "filiais_adicionadas": anterior.get("filiais_adicionadas", []),
                "fonte": str(values[11] or "").strip(),
            }
        )
    result["eventos"].sort(key=lambda e: _date_key(e.get("data", "")))
    return result


MOV_CORES = {
    "entrada": "#235472",
    "saida": "#DC4405",
    "cessao": "#1C6E8C",
    "imovel": "#8A6A45",
    "imovel-saida": "#DC4405",
    "admin": "#1C6E8C",
    "capital-up": "#2F6B3A",
    "capital-down": "#DC4405",
    "sede": "#9A7B2F",
    "objeto": "#9A7B2F",
    "filial": "#77636A",
    "base": "#77817A",
}


def _share(socio: dict) -> str:
    return str(socio.get("participacao", "") or socio.get("quotas", "")).strip()


def _socios_map(event: dict | None) -> dict:
    result = {}
    for socio in ((event or {}).get("socios_apos") or []):
        name = str(socio.get("nome", "")).strip()
        if name:
            result[name] = _share(socio)
    return result


def _valor_num(value: str):
    raw = re.sub(r"[^\d,.-]", "", str(value or ""))
    if not raw:
        return None
    raw = raw.replace(".", "").replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return None


def _moeda(value: float) -> str:
    texto = f"{value:,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    return f"R$ {texto}"


def _quadro_resumo(event: dict) -> str:
    partes = []
    for nome, pct in _socios_map(event).items():
        partes.append(f"{nome} {pct}" if pct else nome)
    return " · ".join(partes) or "Não informado"


def movimentos_do_ato(event: dict, prev: dict | None) -> list[dict]:
    """Movimentações objetivas do ato, com nomes, matrículas e valores."""
    movs: list[dict] = []
    atual = _socios_map(event)
    anterior = _socios_map(prev) if prev else {}

    if prev:
        for nome, pct in atual.items():
            if nome not in anterior:
                movs.append(
                    {
                        "tipo": "entrada",
                        "rotulo": "Entrada de sócio",
                        "nome": nome,
                        "extra": f"{pct} do capital" if pct else "",
                    }
                )
        for nome, pct in anterior.items():
            if nome not in atual:
                movs.append(
                    {
                        "tipo": "saida",
                        "rotulo": "Saída de sócio",
                        "nome": nome,
                        "extra": f"detinha {pct}" if pct else "",
                    }
                )

    for cessao in event.get("cessoes") or []:
        cedente = str(cessao.get("cedente", "")).strip()
        cessionario = str(cessao.get("cessionario", "")).strip()
        if not (cedente or cessionario):
            continue
        extra = " · ".join(
            x
            for x in [
                str(cessao.get("participacao", "")).strip(),
                str(cessao.get("valor", "")).strip(),
                str(cessao.get("observacao", "")).strip(),
            ]
            if x
        )
        movs.append(
            {
                "tipo": "cessao",
                "rotulo": "Cessão de quotas",
                "nome": f"{cedente or '—'} → {cessionario or '—'}",
                "extra": extra,
            }
        )

    for imovel in event.get("imoveis") or []:
        matricula = str(imovel.get("matricula", "")).strip()
        cartorio = str(imovel.get("cartorio", "")).strip() or str(imovel.get("cidade", "")).strip()
        movimento = _normalize(imovel.get("movimento", ""))
        if "said" in movimento or "transfer" in movimento or "restitu" in movimento or "cind" in movimento:
            tipo, rotulo = "imovel-saida", "Saída de imóvel"
        elif "aquis" in movimento or "compr" in movimento:
            tipo, rotulo = "imovel", "Aquisição de imóvel"
        else:
            tipo, rotulo = "imovel", "Integralização de imóvel"
        nome = f"Matrícula {matricula}" if matricula else "Imóvel"
        if cartorio:
            nome = f"{nome} · {cartorio}"
        extra = " — ".join(
            x
            for x in [str(imovel.get("valor", "")).strip(), str(imovel.get("descricao", "")).strip()]
            if x
        )
        movs.append({"tipo": tipo, "rotulo": rotulo, "nome": nome, "extra": extra})

    admin_atual = _joined_people(event.get("administradores_apos", []))
    admin_anterior = _joined_people((prev or {}).get("administradores_apos", [])) if prev else ""
    if prev and admin_atual and admin_atual != admin_anterior:
        movs.append(
            {
                "tipo": "admin",
                "rotulo": "Troca de administração",
                "nome": f"{admin_anterior or '—'} → {admin_atual}",
                "extra": "",
            }
        )

    capital = str(event.get("capital_social_apos", "")).strip()
    capital_anterior = str(event.get("capital_social_anterior", "")).strip()
    if prev and not capital_anterior:
        capital_anterior = str(prev.get("capital_social_apos", "")).strip()
    if capital and capital_anterior and capital != capital_anterior:
        antes, depois = _valor_num(capital_anterior), _valor_num(capital)
        subiu = antes is not None and depois is not None and depois > antes
        delta = ""
        if antes is not None and depois is not None:
            delta = ("+ " if subiu else "− ") + _moeda(abs(depois - antes))
        movs.append(
            {
                "tipo": "capital-up" if subiu else "capital-down",
                "rotulo": "Aumento de capital" if subiu else "Redução de capital",
                "nome": f"{capital_anterior} → {capital}",
                "extra": delta,
            }
        )

    sede = str(event.get("sede_apos", "")).strip()
    sede_anterior = str((prev or {}).get("sede_apos", "")).strip() if prev else ""
    if sede and sede_anterior and _normalize(sede) != _normalize(sede_anterior):
        movs.append(
            {"tipo": "sede", "rotulo": "Mudança de sede", "nome": f"{sede_anterior} → {sede}", "extra": ""}
        )

    objeto = str(event.get("objeto_apos", "")).strip()
    objeto_anterior = str((prev or {}).get("objeto_apos", "")).strip() if prev else ""
    if objeto and objeto_anterior and _normalize(objeto) != _normalize(objeto_anterior):
        movs.append(
            {"tipo": "objeto", "rotulo": "Alteração do objeto social", "nome": objeto, "extra": ""}
        )

    filiais = _joined_filiais(event.get("filiais_adicionadas") or [])
    if filiais:
        movs.append({"tipo": "filial", "rotulo": "Filial aberta", "nome": filiais, "extra": ""})

    if not movs:
        movs.append(
            {
                "tipo": "base",
                "rotulo": "Constituição" if not prev else "Sem alteração relevante",
                "nome": _quadro_resumo(event),
                "extra": (f"Capital {capital}" if capital else ""),
            }
        )
    return movs


def _periodo(events: list[dict]) -> str:
    anos = [str(e.get("data", ""))[-4:] for e in events if str(e.get("data", ""))[-4:].isdigit()]
    if not anos:
        return ""
    return anos[0] if anos[0] == anos[-1] else f"{anos[0]}–{anos[-1]}"


def _detalhe_modal(index: int, event: dict) -> str:
    quadro = "".join(
        f'<span class="tl2-chip">{html.escape(nome)}'
        + (f' <b>{html.escape(pct)}</b>' if pct else "")
        + "</span>"
        for nome, pct in _socios_map(event).items()
    )
    linhas = [
        ("Administração", _joined_people(event.get("administradores_apos", [])) or "—"),
        ("Capital social", str(event.get("capital_social_apos", "")) or "—"),
        ("Sede", str(event.get("sede_apos", "")) or "—"),
        ("Objeto social", str(event.get("objeto_apos", "")) or "—"),
        ("Filiais", _joined_filiais(event.get("filiais_apos") or []) or "—"),
    ]
    imoveis = "".join(
        f'<li><b>Matrícula {html.escape(str(i.get("matricula", "")) or "—")}</b> — '
        f'{html.escape(str(i.get("cartorio", "")) or str(i.get("cidade", "")))} · '
        f'{html.escape(str(i.get("valor", "")) or "valor não informado")} · '
        f'{html.escape(str(i.get("movimento", "")) or "—")}</li>'
        for i in (event.get("imoveis") or [])
    )
    cessoes = "".join(
        f'<li><b>{html.escape(str(c.get("cedente", "")) or "—")} → '
        f'{html.escape(str(c.get("cessionario", "")) or "—")}</b> — '
        + html.escape(
            " · ".join(
                x
                for x in [
                    str(c.get("participacao", "")).strip(),
                    str(c.get("valor", "")).strip(),
                    str(c.get("observacao", "")).strip(),
                ]
                if x
            )
            or "sem valores informados"
        )
        + "</li>"
        for c in (event.get("cessoes") or [])
    )
    return f"""
      <input type="checkbox" id="tl2-det-{index}" class="tl2-toggle">
      <label class="tl2-btn" for="tl2-det-{index}">Ver detalhamento</label>
      <div class="tl2-modal">
        <label class="tl2-backdrop" for="tl2-det-{index}"></label>
        <div class="tl2-modal-card">
          <div class="tl2-modal-head">
            <div>
              <span class="tl2-modal-date">{html.escape(str(event.get("data", "")) or "—")}</span>
              <strong>{html.escape(str(event.get("ato", "")) or "Ato societário")}</strong>
              <small>{html.escape(str(event.get("numero_arquivamento", "")))}</small>
            </div>
            <label class="tl2-close" for="tl2-det-{index}">Fechar</label>
          </div>
          <div class="tl2-modal-body">
            <h4>Detalhamento do ato</h4>
            <p>{html.escape(str(event.get("detalhamento", "")) or "Não informado.")}</p>
            <h4>Quadro societário após o ato</h4>
            <div class="tl2-chips">{quadro or '<span class="tl2-chip">Não informado</span>'}</div>
            {'<h4>Transmissão de quotas</h4><ul class="tl2-list">' + cessoes + '</ul>' if cessoes else ''}
            {'<h4>Bens imóveis</h4><ul class="tl2-list">' + imoveis + '</ul>' if imoveis else ''}
            <h4>Estado da sociedade</h4>
            <dl class="tl2-dl">
              {''.join(f'<dt>{html.escape(k)}</dt><dd>{html.escape(str(v))}</dd>' for k, v in linhas)}
            </dl>
            <p class="tl2-fonte">Fonte: {html.escape(str(event.get("fonte", "")) or "não informada")}</p>
          </div>
        </div>
      </div>
    """


def render_timeline_html(data: dict) -> str:
    events = data.get("eventos", []) or []
    if not events:
        return '<div class="timeline-empty">A análise ainda não gerou eventos societários.</div>'

    colunas = []
    for index, event in enumerate(events):
        prev = events[index - 1] if index else None
        blocos = "".join(
            f'<div class="tl2-move tl2-move-{m["tipo"]}">'
            f'<span class="tl2-move-label">{html.escape(m["rotulo"])}</span>'
            f'<strong>{html.escape(m["nome"] or "—")}</strong>'
            + (f'<small>{html.escape(m["extra"])}</small>' if m.get("extra") else "")
            + "</div>"
            for m in movimentos_do_ato(event, prev)
        )
        colunas.append(
            f"""
        <article class="tl2-col">
          <div class="tl2-col-head">
            <strong>{html.escape(str(event.get("data", "")) or "—")}</strong>
            <span>{html.escape(str(event.get("ato", "")) or "Ato societário")}</span>
          </div>
          <div class="tl2-axis"><i class="tl2-pin"></i></div>
          <div class="tl2-moves">{blocos}</div>
          <div class="tl2-quadro">Quadro após o ato<strong>{html.escape(_quadro_resumo(event))}</strong></div>
          {_detalhe_modal(index, event)}
        </article>
        """
        )

    return f"""
    <section class="tl2-shell" id="timeline-export-area">
      <div class="tl2-head">
        <span class="tl2-kicker">Cronologia societária</span>
        <h2>{html.escape(data.get("empresa", "") or "Sociedade analisada")}</h2>
        <p>{("CNPJ " + html.escape(data.get("cnpj", "")) + " · ") if data.get("cnpj") else ""}
           {len(events)} atos · {_periodo(events)}</p>
      </div>
      <div class="tl2-scroll">
        <div class="tl2-track" style="--tl2-cols:{len(events)}">{"".join(colunas)}</div>
      </div>
    </section>
    """


def timeline_analisar(files):
    yield (
        "Preparando os atos societários...",
        '<div class="timeline-loading">Lendo os documentos e reconstruindo a evolução societária…</div>',
        [],
        {},
    )
    data = _extract(files)
    yield (
        f"Concluído: {len(data.get('eventos', []))} evento(s) societário(s) identificado(s).",
        render_timeline_html(data),
        data_to_rows(data),
        data,
    )


def timeline_toggle_edicao(editando: bool, rows: Any, data: dict):
    if not editando:
        return (
            True,
            gr.update(value="Concluir edição"),
            gr.update(visible=True, interactive=True, value=data_to_rows(data or {})),
            gr.update(),
            data,
            "Modo de edição ativo. Corrija os campos na tabela abaixo.",
        )
    updated = rows_to_data(rows, data or {})
    return (
        False,
        gr.update(value="Editar"),
        gr.update(visible=False, value=data_to_rows(updated)),
        render_timeline_html(updated),
        updated,
        "Edições salvas e visualização atualizada.",
    )


def timeline_ver_tabela(data: dict):
    events = (data or {}).get("eventos", [])
    rows = [
        [
            event.get("data", ""),
            event.get("ato", ""),
            " ".join(detalhamento_word(event, events[index - 1] if index else None)),
        ]
        for index, event in enumerate(events)
    ]
    return gr.update(visible=True, value=rows)


def _font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size)
    return ImageFont.load_default()


def _wrap(draw: ImageDraw.ImageDraw, text: str, font, max_width: int, max_lines: int = 4) -> list[str]:
    words = str(text or "").split()
    if not words:
        return []
    lines, current = [], ""
    for word in words:
        trial = f"{current} {word}".strip()
        if draw.textbbox((0, 0), trial, font=font)[2] <= max_width or not current:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines[:max_lines]


def _hex(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def _tint(value: str, alpha: float = 0.09) -> tuple[int, int, int]:
    r, g, b = _hex(value)
    return (
        int(r * alpha + 255 * (1 - alpha)),
        int(g * alpha + 255 * (1 - alpha)),
        int(b * alpha + 255 * (1 - alpha)),
    )


def timeline_exportar_imagem(data: dict) -> str:
    """Cronologia enxuta, em alta resolução, para colar no dossiê."""
    events = (data or {}).get("eventos", [])
    if not events:
        raise gr.Error("Gere a timeline antes de exportar.")

    scale = 2
    col_w = 250 * scale
    gap = 14 * scale
    pad = 34 * scale
    header_h = 128 * scale
    axis_h = 46 * scale
    inner_w = col_w - 26 * scale

    text, muted, line = "#2C302C", "#737670", "#E4E4E0"

    probe = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(probe)
    f_titulo = _font(19 * scale, True)
    f_meta = _font(11 * scale)
    f_data = _font(13 * scale, True)
    f_ato = _font(9 * scale, True)
    f_rotulo = _font(8 * scale, True)
    f_nome = _font(10 * scale, True)
    f_extra = _font(8 * scale)
    f_quadro = _font(9 * scale)

    colunas = []
    for index, event in enumerate(events):
        prev = events[index - 1] if index else None
        blocos = []
        altura = 0
        for mov in movimentos_do_ato(event, prev):
            nome_linhas = _wrap(draw, mov["nome"], f_nome, inner_w - 16 * scale, 4)
            extra_linhas = _wrap(draw, mov.get("extra", ""), f_extra, inner_w - 16 * scale, 3)
            bloco_h = (
                9 * scale
                + 12 * scale
                + len(nome_linhas) * 14 * scale
                + (len(extra_linhas) * 12 * scale + 3 * scale if extra_linhas else 0)
                + 9 * scale
            )
            blocos.append(
                {
                    "cor": MOV_CORES.get(mov["tipo"], "#77817A"),
                    "rotulo": mov["rotulo"].upper(),
                    "nome": nome_linhas,
                    "extra": extra_linhas,
                    "altura": bloco_h,
                }
            )
            altura += bloco_h + 8 * scale
        quadro_linhas = _wrap(draw, _quadro_resumo(event), f_quadro, inner_w, 4)
        altura += 16 * scale + len(quadro_linhas) * 13 * scale
        colunas.append({"event": event, "blocos": blocos, "quadro": quadro_linhas, "altura": altura})

    corpo_h = max(c["altura"] for c in colunas)
    width = pad * 2 + len(events) * col_w + (len(events) - 1) * gap
    height = header_h + axis_h + corpo_h + pad
    image = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(image)

    draw.text((pad, 30 * scale), "CRONOLOGIA SOCIETÁRIA", fill="#DC4405", font=_font(9 * scale, True))
    draw.text((pad, 46 * scale), str(data.get("empresa", "") or "Sociedade analisada"), fill="#1F211F", font=f_titulo)
    meta = " · ".join(
        x
        for x in [
            ("CNPJ " + str(data.get("cnpj", ""))) if data.get("cnpj") else "",
            f"{len(events)} atos",
            _periodo(events),
        ]
        if x
    )
    draw.text((pad, 78 * scale), meta, fill=muted, font=f_meta)
    draw.line((pad, header_h - 22 * scale, width - pad, header_h - 22 * scale), fill="#1F211F", width=2 * scale)

    axis_y = header_h + axis_h // 2
    draw.line((pad, axis_y, width - pad, axis_y), fill=line, width=2 * scale)

    for index, coluna in enumerate(colunas):
        x = pad + index * (col_w + gap)
        event = coluna["event"]
        draw.text((x, header_h - 4 * scale), str(event.get("data", "") or "—"), fill=text, font=f_data)
        draw.text(
            (x, header_h + 16 * scale),
            str(event.get("ato", "") or "ATO SOCIETÁRIO").upper(),
            fill=muted,
            font=f_ato,
        )
        raio = 6 * scale
        draw.ellipse((x - raio, axis_y - raio, x + raio, axis_y + raio), fill="#DC4405")

        y = header_h + axis_h
        for bloco in coluna["blocos"]:
            draw.rectangle((x, y, x + col_w, y + bloco["altura"]), fill=_tint(bloco["cor"]))
            draw.rectangle((x, y, x + 3 * scale, y + bloco["altura"]), fill=bloco["cor"])
            ty = y + 9 * scale
            draw.text((x + 12 * scale, ty), bloco["rotulo"], fill=bloco["cor"], font=f_rotulo)
            ty += 14 * scale
            for linha_texto in bloco["nome"]:
                draw.text((x + 12 * scale, ty), linha_texto, fill="#1F211F", font=f_nome)
                ty += 14 * scale
            if bloco["extra"]:
                ty += 2 * scale
                for linha_texto in bloco["extra"]:
                    draw.text((x + 12 * scale, ty), linha_texto, fill=muted, font=f_extra)
                    ty += 12 * scale
            y += bloco["altura"] + 8 * scale

        y += 6 * scale
        draw.line((x, y, x + col_w, y), fill=line, width=1 * scale)
        y += 8 * scale
        draw.text((x, y), "QUADRO APÓS O ATO", fill="#A9ABA6", font=f_rotulo)
        y += 13 * scale
        for linha_texto in coluna["quadro"]:
            draw.text((x, y), linha_texto, fill="#3F423E", font=f_quadro)
            y += 13 * scale

    output = Path(tempfile.gettempdir()) / "timeline_societaria_alta_resolucao.png"
    image.save(output, format="PNG", optimize=True, dpi=(192, 192))
    return str(output)



def _medir_timeline_vertical(events: list[dict], scale: float, width: int):
    pad = 34 * scale
    axis_x = pad + 6 * scale
    content_x = axis_x + 26 * scale
    content_w = width - content_x - pad
    header_h = 118 * scale
    gap_evento = 22 * scale

    probe = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(probe)
    fontes = {
        "kicker": _font(round(9 * scale), True),
        "titulo": _font(round(19 * scale), True),
        "meta": _font(round(11 * scale)),
        "data": _font(round(13 * scale), True),
        "ato": _font(round(9 * scale), True),
        "rotulo": _font(round(8 * scale), True),
        "nome": _font(round(10 * scale), True),
        "extra": _font(round(8 * scale)),
        "quadro": _font(round(9 * scale)),
    }

    itens = []
    for idx, event in enumerate(events):
        prev = events[idx - 1] if idx else None
        blocos = []
        for mov in movimentos_do_ato(event, prev):
            nome_linhas = _wrap(draw, mov["nome"], fontes["nome"], content_w - 16 * scale, 4)
            extra_linhas = _wrap(draw, mov.get("extra", ""), fontes["extra"], content_w - 16 * scale, 3)
            bloco_h = (
                9 * scale
                + 12 * scale
                + len(nome_linhas) * 14 * scale
                + (len(extra_linhas) * 12 * scale + 3 * scale if extra_linhas else 0)
                + 9 * scale
            )
            blocos.append(
                {
                    "cor": MOV_CORES.get(mov["tipo"], "#77817A"),
                    "rotulo": mov["rotulo"].upper(),
                    "nome": nome_linhas,
                    "extra": extra_linhas,
                    "altura": bloco_h,
                }
            )
        quadro_linhas = _wrap(draw, _quadro_resumo(event), fontes["quadro"], content_w, 4)
        altura = 20 * scale + sum(b["altura"] + 8 * scale for b in blocos) + 16 * scale + len(quadro_linhas) * 13 * scale
        itens.append({"event": event, "blocos": blocos, "quadro": quadro_linhas, "altura": altura})

    body_h = sum(item["altura"] for item in itens) + gap_evento * max(0, len(itens) - 1)
    height = header_h + body_h + pad
    return {
        "scale": scale, "width": width, "height": height, "pad": pad, "axis_x": axis_x,
        "content_x": content_x, "content_w": content_w, "header_h": header_h,
        "gap_evento": gap_evento, "body_h": body_h, "itens": itens, "fontes": fontes,
    }


def timeline_exportar_imagem_vertical(data: dict) -> str:
    """Cronologia em layout vertical (cards coloridos, mesma linguagem visual da versão horizontal),
    dimensionada para caber em uma página A4 vertical do Word."""
    events = (data or {}).get("eventos", [])
    if not events:
        raise gr.Error("Gere a timeline antes de exportar.")

    width = 2000
    info = _medir_timeline_vertical(events, 2.0, width)
    a4_ratio = 297 / 210  # altura/largura de uma folha A4 vertical
    limite = width * a4_ratio
    if info["height"] > limite:
        fator = max(0.45, limite / info["height"])
        info = _medir_timeline_vertical(events, 2.0 * fator, width)

    scale = info["scale"]
    pad, axis_x, content_x, content_w = info["pad"], info["axis_x"], info["content_x"], info["content_w"]
    header_h, gap_evento, body_h, itens, fontes = (
        info["header_h"], info["gap_evento"], info["body_h"], info["itens"], info["fontes"]
    )
    f_kicker, f_titulo, f_meta = fontes["kicker"], fontes["titulo"], fontes["meta"]
    f_data, f_ato = fontes["data"], fontes["ato"]
    f_rotulo, f_nome, f_extra, f_quadro = fontes["rotulo"], fontes["nome"], fontes["extra"], fontes["quadro"]

    text, muted, line, azul = "#2C302C", "#737670", "#E4E4E0", "#1A56A0"

    image = Image.new("RGB", (width, round(info["height"])), "#FFFFFF")
    draw = ImageDraw.Draw(image)

    draw.text((pad, 30 * scale), "CRONOLOGIA SOCIETÁRIA", fill=azul, font=f_kicker)
    draw.text((pad, 46 * scale), str(data.get("empresa", "") or "Sociedade analisada"), fill="#1F211F", font=f_titulo)
    meta = " · ".join(
        x
        for x in [
            ("CNPJ " + str(data.get("cnpj", ""))) if data.get("cnpj") else "",
            f"{len(events)} atos",
            _periodo(events),
        ]
        if x
    )
    draw.text((pad, 78 * scale), meta, fill=muted, font=f_meta)
    draw.line((pad, header_h - 22 * scale, width - pad, header_h - 22 * scale), fill="#1F211F", width=max(1, round(2 * scale)))
    draw.line((axis_x, header_h, axis_x, header_h + body_h), fill=line, width=max(1, round(2 * scale)))

    y = header_h
    for item in itens:
        event = item["event"]
        r = 6 * scale
        draw.ellipse((axis_x - r, y + 10 * scale - r, axis_x + r, y + 10 * scale + r), fill=azul)

        draw.text((content_x, y), str(event.get("data", "") or "—"), fill=text, font=f_data)
        data_w = draw.textbbox((0, 0), str(event.get("data", "") or "—"), font=f_data)[2]
        draw.text(
            (content_x + data_w + 10 * scale, y + 2 * scale),
            str(event.get("ato", "") or "Ato societário").upper(),
            fill=muted,
            font=f_ato,
        )
        yy = y + 20 * scale

        for bloco in item["blocos"]:
            draw.rectangle((content_x, yy, content_x + content_w, yy + bloco["altura"]), fill=_tint(bloco["cor"]))
            draw.rectangle((content_x, yy, content_x + 3 * scale, yy + bloco["altura"]), fill=bloco["cor"])
            ty = yy + 9 * scale
            draw.text((content_x + 12 * scale, ty), bloco["rotulo"], fill=bloco["cor"], font=f_rotulo)
            ty += 14 * scale
            for linha_texto in bloco["nome"]:
                draw.text((content_x + 12 * scale, ty), linha_texto, fill="#1F211F", font=f_nome)
                ty += 14 * scale
            if bloco["extra"]:
                ty += 2 * scale
                for linha_texto in bloco["extra"]:
                    draw.text((content_x + 12 * scale, ty), linha_texto, fill=muted, font=f_extra)
                    ty += 12 * scale
            yy += bloco["altura"] + 8 * scale

        yy += 6 * scale
        for linha_texto in item["quadro"]:
            draw.text((content_x, yy), linha_texto, fill="#3F423E", font=f_quadro)
            yy += 13 * scale

        y += item["altura"] + gap_evento

    output = Path(tempfile.gettempdir()) / "timeline_societaria_vertical.png"
    image.save(output, format="PNG", optimize=True, dpi=(192, 192))
    return str(output)


TL2_CSS = """
.timeline-empty, .timeline-loading {
    margin: 18px 0; padding: 44px 24px; text-align: center; background: #f5f4f1;
    border: 1px dashed #cbc9c1; color: #71746f;
}
.timeline-actions { align-items: flex-start; margin-top: 18px; }
.timeline-action-stack { gap: 8px !important; }

/* ── Timeline Societária — cronologia nomeada (v2) ──────────────────────── */
.tl2-shell {
    margin-top: 18px; background: #fff; border: 1px solid #d9d9d5;
    box-shadow: 0 12px 28px rgba(36, 38, 35, .07);
}
.tl2-head { padding: 24px 28px 20px; border-bottom: 1px solid #dededb; }
.tl2-kicker {
    display: block; color: #1A56A0; font-size: 11px; font-weight: 800; letter-spacing: .14em;
    text-transform: uppercase;
}
.tl2-head h2 { margin: 6px 0 4px; font-size: 25px; font-weight: 700; color: #2c302c; }
.tl2-head p { margin: 0; font-size: 13px; color: #737670; }
.tl2-scroll { overflow-x: auto; padding: 26px 28px 30px; }
.tl2-track {
    display: grid; grid-template-columns: repeat(var(--tl2-cols), 200px);
    min-width: 100%; column-gap: 0;
}
.tl2-col { padding: 0 10px; display: flex; flex-direction: column; }
.tl2-col-head { text-align: center; }
.tl2-col-head strong { display: block; font-size: 14px; font-weight: 800; color: #1f211f; }
.tl2-col-head span {
    display: block; margin-top: 2px; font-size: 10px; font-weight: 800; letter-spacing: .06em;
    text-transform: uppercase; color: #737670;
}
.tl2-axis {
    margin-top: 12px; height: 3px; background: #e4e4e0; display: flex; align-items: center;
    justify-content: center;
}
.tl2-pin {
    width: 13px; height: 13px; border-radius: 50%; background: #1A56A0; border: 2px solid #fff;
    box-shadow: 0 0 0 1px #1A56A0;
}
.tl2-moves { margin-top: 18px; display: flex; flex-direction: column; gap: 7px; }
.tl2-move { padding: 8px 9px; border-left: 3px solid #77817a; background: #f6f6f4; }
.tl2-move-label {
    display: block; font-size: 9px; font-weight: 800; letter-spacing: .08em; text-transform: uppercase;
    color: #77817a;
}
.tl2-move strong {
    display: block; margin-top: 3px; font-size: 12px; font-weight: 700; color: #1f211f;
    line-height: 1.35; text-wrap: pretty;
}
.tl2-move small { display: block; margin-top: 3px; font-size: 10px; color: #666963; line-height: 1.4; }
.tl2-move-entrada { border-left-color: #235472; background: rgba(35, 84, 114, .07); }
.tl2-move-entrada .tl2-move-label { color: #235472; }
.tl2-move-saida { border-left-color: #dc4405; background: rgba(220, 68, 5, .07); }
.tl2-move-saida .tl2-move-label { color: #dc4405; }
.tl2-move-cessao { border-left-color: #1c6e8c; background: rgba(28, 110, 140, .07); }
.tl2-move-cessao .tl2-move-label { color: #1c6e8c; }
.tl2-move-imovel { border-left-color: #8a6a45; background: rgba(138, 106, 69, .08); }
.tl2-move-imovel .tl2-move-label { color: #8a6a45; }
.tl2-move-imovel-saida { border-left-color: #dc4405; background: rgba(220, 68, 5, .07); }
.tl2-move-imovel-saida .tl2-move-label { color: #dc4405; }
.tl2-move-admin { border-left-color: #1c6e8c; background: rgba(28, 110, 140, .07); }
.tl2-move-admin .tl2-move-label { color: #1c6e8c; }
.tl2-move-capital-up { border-left-color: #2f6b3a; background: rgba(47, 107, 58, .07); }
.tl2-move-capital-up .tl2-move-label { color: #2f6b3a; }
.tl2-move-capital-down { border-left-color: #dc4405; background: rgba(220, 68, 5, .07); }
.tl2-move-capital-down .tl2-move-label { color: #dc4405; }
.tl2-move-sede, .tl2-move-objeto { border-left-color: #9a7b2f; background: rgba(154, 123, 47, .08); }
.tl2-move-sede .tl2-move-label, .tl2-move-objeto .tl2-move-label { color: #9a7b2f; }
.tl2-move-filial { border-left-color: #77636a; background: rgba(119, 99, 106, .08); }
.tl2-move-filial .tl2-move-label { color: #77636a; }
.tl2-quadro {
    margin-top: 12px; padding-top: 10px; border-top: 1px dashed #dededa; font-size: 10px;
    color: #737670; line-height: 1.45;
}
.tl2-quadro strong { display: block; margin-top: 2px; color: #3f423e; font-weight: 700; }
.tl2-toggle { display: none !important; }
.tl2-btn {
    margin-top: 10px; padding: 7px 10px; background: #fff; border: 1px solid #dcdcd7;
    border-radius: 999px; font-size: 10px; font-weight: 800; letter-spacing: .06em;
    text-transform: uppercase; color: #4a4d48; text-align: center; cursor: pointer;
    transition: background .16s ease, color .16s ease;
}
.tl2-btn:hover { background: #f5f5f2; color: #1f211f; }
.tl2-modal {
    display: none; position: fixed; inset: 0; z-index: 90; align-items: center;
    justify-content: center; padding: 40px;
}
.tl2-toggle:checked ~ .tl2-modal { display: flex; }
.tl2-backdrop { position: absolute; inset: 0; background: rgba(31, 33, 31, .42); cursor: pointer; }
.tl2-modal-card {
    position: relative; width: 640px; max-height: 84vh; overflow-y: auto; background: #fff;
    border: 1px solid #d9d9d5; box-shadow: 0 24px 60px rgba(31, 33, 31, .28); text-align: left;
}
.tl2-modal-head {
    display: flex; align-items: flex-start; justify-content: space-between; gap: 16px;
    padding: 20px 24px 16px; border-bottom: 1px solid #ededea;
}
.tl2-modal-date {
    display: block; font-size: 11px; font-weight: 800; letter-spacing: .1em; text-transform: uppercase;
    color: #dc4405;
}
.tl2-modal-head strong { display: block; margin-top: 5px; font-size: 19px; font-weight: 800; color: #1f211f; }
.tl2-modal-head small { display: block; margin-top: 2px; font-size: 11px; color: #92948f; }
.tl2-close {
    flex: none; padding: 7px 14px; background: #f5f5f2; border: 1px solid #e4e4e0; border-radius: 999px;
    font-size: 10px; font-weight: 800; letter-spacing: .06em; text-transform: uppercase;
    color: #4a4d48; cursor: pointer;
}
.tl2-close:hover { background: #ededea; }
.tl2-modal-body { padding: 20px 24px 24px; }
.tl2-modal-body h4 {
    margin: 18px 0 8px; font-size: 10px; font-weight: 800; letter-spacing: .1em;
    text-transform: uppercase; color: #92948f;
}
.tl2-modal-body h4:first-child { margin-top: 0; }
.tl2-modal-body p { margin: 0; font-size: 13px; line-height: 1.65; color: #3f423e; text-wrap: pretty; }
.tl2-chips { display: flex; flex-wrap: wrap; gap: 8px; }
.tl2-chip {
    padding: 5px 10px; background: #f4f7f9; border: 1px solid #e2eaef; font-size: 12px; color: #235472;
}
.tl2-list { margin: 0; padding-left: 18px; font-size: 12px; line-height: 1.65; color: #3f423e; }
.tl2-dl { display: grid; grid-template-columns: auto 1fr; gap: 6px 14px; margin: 0; font-size: 12px; }
.tl2-dl dt { color: #92948f; }
.tl2-dl dd { margin: 0; color: #4a4d48; font-weight: 600; }
.tl2-fonte { margin-top: 18px !important; font-size: 10px !important; color: #92948f !important; }
"""


def timeline_exportar_html(data: dict) -> str:
    """Exporta a timeline interativa (colunas + modais de detalhamento) como um HTML autocontido."""
    events = (data or {}).get("eventos", [])
    if not events:
        raise gr.Error("Gere a timeline antes de exportar.")
    empresa = html.escape(str(data.get("empresa", "") or "Sociedade analisada"))
    corpo = render_timeline_html(data)
    documento = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Cronologia Societária — {empresa}</title>
<style>
body {{ margin: 0; padding: 0; background: #FFFFFF; font-family: Arial, Helvetica, sans-serif; }}
{TL2_CSS}
</style>
</head>
<body>
{corpo}
</body>
</html>
"""
    output = Path(tempfile.gettempdir()) / "timeline_societaria_interativa.html"
    output.write_text(documento, encoding="utf-8")
    return str(output)


def _normalize(value: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", str(value or "").lower()) if unicodedata.category(c) != "Mn"
    )


def _find_timeline_table(doc: Document):
    body = doc.element.body
    for paragraph in doc.paragraphs:
        normalized = _normalize(paragraph.text)
        if "cronologia societaria" not in normalized:
            continue
        start = list(body).index(paragraph._p)
        for child in list(body)[start + 1 :]:
            if child.tag == qn("w:tbl"):
                from docx.table import Table

                return Table(child, doc)
            if child.tag == qn("w:p") and "evidencias" in _normalize("".join(child.itertext())):
                break
    return None


def _set_cell(cell, value: str, bold: bool = False, color: str = "555555"):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(str(value or ""))
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(color)
    run.bold = bold


def detalhamento_word(event: dict, prev: dict | None) -> list[str]:
    """Tópicos do dossiê: detalhamento do ato + movimentações objetivas, sem a fonte."""
    itens = []
    detalhe = str(event.get("detalhamento", "")).strip()
    if detalhe:
        itens.append(detalhe)
    for m in movimentos_do_ato(event, prev):
        if m["tipo"] == "base":
            continue
        itens.append(f"{m['rotulo']}: {m['nome']}" + (f" ({m['extra']})" if m.get("extra") else ""))
    quadro = _quadro_resumo(event)
    if quadro and quadro != "Não informado":
        itens.append("Quadro societário após o ato: " + quadro)
    admin = _joined_people(event.get("administradores_apos", []))
    if admin:
        itens.append("Administração: " + admin)
    capital = str(event.get("capital_social_apos", "")).strip()
    if capital:
        itens.append("Capital social: " + capital)
    return itens or ["Sem alterações relevantes registradas."]


def _set_cell_bullets(cell, itens: list[str], color: str = "555555"):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for index, item in enumerate(itens):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(3)
        run = paragraph.add_run("• " + str(item))
        run.font.name = "Arial"
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(color)


def _set_col_widths(table, widths=(Inches(1.0), Inches(1.1), Inches(4.4))):
    """Data/Ato mais estreitos; Detalhamento absorve o espaço restante."""
    table.autofit = False
    for row in table.rows:
        for index, width in enumerate(widths):
            if index < len(row.cells):
                row.cells[index].width = width
    for index, width in enumerate(widths):
        if index < len(table.columns):
            table.columns[index].width = width


def _fill_word_table(table, data: dict):
    if not table.rows:
        return
    template_row = copy.deepcopy(table.rows[1]._tr if len(table.rows) > 1 else table.rows[0]._tr)
    while len(table.rows) > 1:
        table._tbl.remove(table.rows[-1]._tr)
    events = (data or {}).get("eventos", [])
    if not events:
        table._tbl.append(copy.deepcopy(template_row))
        _set_col_widths(table)
        return
    for index, event in enumerate(events):
        row_xml = copy.deepcopy(template_row)
        table._tbl.append(row_xml)
        row = table.rows[-1]
        prev = events[index - 1] if index else None
        _set_cell(row.cells[0], event.get("data", ""))
        _set_cell(row.cells[1], event.get("ato", ""))
        _set_cell_bullets(row.cells[2], detalhamento_word(event, prev))
    _set_col_widths(table)


def timeline_gerar_word(data: dict, dossie_file: Any = None) -> str:
    if not (data or {}).get("eventos"):
        raise gr.Error("Gere a timeline antes de exportar a tabela Word.")
    dossier_path = _file_path(dossie_file)
    if dossier_path and Path(dossier_path).exists():
        doc = Document(dossier_path)
        output_name = "dossie_com_cronologia_societaria.docx"
    elif TEMPLATE_PATH.exists():
        doc = Document(str(TEMPLATE_PATH))
        output_name = "cronologia_societaria.docx"
    else:
        doc = Document()
        title = doc.add_paragraph()
        run = title.add_run("c. Cronologia Societária — Atos Relevantes")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        table = doc.add_table(rows=2, cols=3)
        table.style = "Table Grid"
        for cell, label in zip(table.rows[0].cells, ["DATA", "ATO", "DETALHAMENTO"]):
            _set_cell(cell, label, bold=True, color="FFFFFF")
        output_name = "cronologia_societaria.docx"

    table = _find_timeline_table(doc)
    if table is None:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run("c. Cronologia Societária — Atos Relevantes")
        run.bold = True
        run.font.name = "Arial"
        run.font.size = Pt(10)
        table = doc.add_table(rows=2, cols=3)
        table.style = "Table Grid"
        for cell, label in zip(table.rows[0].cells, ["DATA", "ATO", "DETALHAMENTO"]):
            _set_cell(cell, label, bold=True, color="FFFFFF")
    _fill_word_table(table, data)

    output = Path(tempfile.gettempdir()) / output_name
    doc.save(output)
    return str(output)
