# -*- coding: utf-8 -*-
import os
import re
import time
import tempfile
from pathlib import Path

import fitz
from google import genai
from google.genai import types
from docx import Document as DocxDoc
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# Timeout (ms) por requisicao HTTP ao Gemini. Converte travamento de socket
# (conexao viva mas sem resposta) em excecao, permitindo que o _retry atue.
# 10 min cobre folgado ate a consolidacao; acima disso e hang de verdade.
GEMINI_TIMEOUT_MS = int(os.getenv("GEMINI_TIMEOUT_MS", "600000"))


def _get_clients_rj():
    k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
    k2 = os.getenv("GEMINI_API_KEY_2") or k1
    if not k1:
        raise RuntimeError("GEMINI_API_KEY_1 não configurada.")
    http_opts = types.HttpOptions(timeout=GEMINI_TIMEOUT_MS)
    return (
        genai.Client(api_key=k1, http_options=http_opts),
        genai.Client(api_key=k2, http_options=http_opts),
    )


def _get_clients_proc():
    k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
    k2 = os.getenv("GEMINI_API_KEY_2") or k1
    if not k1:
        raise RuntimeError("GEMINI_API_KEY_1 não configurada.")
    http_opts = types.HttpOptions(timeout=GEMINI_TIMEOUT_MS)
    return (
        genai.Client(api_key=k1, http_options=http_opts),
        genai.Client(api_key=k2, http_options=http_opts),
    )

def _retry(fn, tentativas=5, espera_base=20):
    for t in range(1, tentativas + 1):
        try:
            return fn()
        except Exception as e:
            msg = str(e)
            msg_low = msg.lower()
            retryable = (
                any(c in msg for c in ["503", "500", "504", "UNAVAILABLE", "overloaded"])
                or any(c in msg_low for c in [
                    "timeout", "timed out", "deadline", "deadlineexceeded",
                    "connection reset", "connection error", "connecterror",
                    "read timed out", "readtimeout",
                ])
            )
            if retryable and t < tentativas:
                time.sleep(espera_base * t)
            else:
                raise
    raise RuntimeError("Falha após todas as tentativas.")


def _gerar_docx(relatorio: str, titulo: str = "Análise Jurídica") -> str:
    caminho = os.path.join(tempfile.gettempdir(), f"{titulo.lower().replace(' ', '_')}.docx")
    doc = DocxDoc()

    estilo = doc.styles["Normal"]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(11)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(titulo)
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
    doc.add_paragraph()

    REF_PAT = re.compile(
        r'\(([^)]*(?:Mov\.|fls\.|Fls\.|Evento\s+n[oº°]?|ID\s*\d+|\|\s*fls\.)[^)]*)\)'
    )
    HEADING_PAT = re.compile(
        r'^[A-Z]\.\s+[A-ZÁÉÍÓÚ]|^[A-Z]\.\d+\.|^\d+\.\s+[A-ZÁÉÍÓÚ]'
    )

    BULLET_MAP = [
        ('•',  'List Bullet',   None),
        ('∘',  'List Bullet 2', None),
        ('▪',  'List Bullet 3', None),
        ('▵',  'List Bullet 3', Pt(12)),
    ]

    def _add_runs(paragraph, text):
        pos = 0
        for m in REF_PAT.finditer(text):
            if m.start() > pos:
                run = paragraph.add_run(text[pos:m.start()])
                run.font.name = "Calibri"
            run = paragraph.add_run(f"({m.group(1)})")
            run.italic = True
            run.font.name = "Calibri"
            pos = m.end()
        if pos < len(text):
            run = paragraph.add_run(text[pos:])
            run.font.name = "Calibri"

    for linha in relatorio.split("\n"):
        s = linha.strip()
        if not s:
            doc.add_paragraph()
            continue

        if HEADING_PAT.match(s):
            p = doc.add_paragraph()
            run = p.add_run(s)
            run.bold = True
            run.font.name = "Calibri"
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
            continue

        bullet_style = None
        extra_indent = None
        for char, style, extra in BULLET_MAP:
            if s.startswith(char):
                bullet_style = style
                extra_indent = extra
                s = s[len(char):].strip()
                break

        if bullet_style:
            try:
                p = doc.add_paragraph(style=bullet_style)
            except Exception:
                p = doc.add_paragraph()
            if extra_indent:
                current = p.paragraph_format.left_indent or Pt(0)
                p.paragraph_format.left_indent = current + extra_indent
        else:
            p = doc.add_paragraph()

        _add_runs(p, s)

    doc.save(caminho)
    return caminho


def _responder_pergunta_generica(pergunta: str, relatorio: str, client, model: str) -> str:
    if not relatorio.strip():
        return "Gere uma análise primeiro."
    if not pergunta.strip():
        return "Digite uma pergunta."
    prompt = (
        f"Com base no relatório jurídico abaixo, responda de forma objetiva e precisa:\n\n"
        f"PERGUNTA: {pergunta.strip()}\n\n"
        f"RELATÓRIO:\n{relatorio[:80_000]}"
    )
    def _fn():
        return client.models.generate_content(model=model, contents=[prompt]).text
    return _retry(_fn)



def _comprimir_pdf_limite(path: str, max_mb: float = 40.0) -> tuple:
    """
    Garante que o PDF resultante fique abaixo de max_mb.
    Tenta compressão normal primeiro; se ainda acima do limite,
    re-renderiza TODAS as páginas (inclusive texto) como JPEG,
    começando em 110 DPI (suficiente para leitura por IA).
    """
    orig_mb = Path(path).stat().st_size / 1_048_576

    # Fast path — já está dentro do limite, não faz nada
    if orig_mb <= max_mb:
        return path, orig_mb, orig_mb

    comp, _, comp_mb = _comprimir_pdf(path)
    if comp_mb <= max_mb:
        return comp, orig_mb, comp_mb

    if comp != path:
        try: os.remove(comp)
        except: pass

    # 110 DPI é o ponto de partida — evita passes desnecessários em 150/130 DPI
    # que para PDFs de texto raramente cabem no limite
    for dpi, quality in [(110, 42), (96, 40), (80, 38)]:
        tmp_path = None
        try:
            doc = fitz.open(path)
            new_doc = fitz.open()
            for page in doc:
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                pix = page.get_pixmap(matrix=mat)
                img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
                rect = page.rect
                new_page = new_doc.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=img_bytes, keep_proportion=False)
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp_path = tmp.name
            tmp.close()
            new_doc.save(tmp_path, garbage=4, deflate=True)
            new_doc.close()
            doc.close()
            result_mb = Path(tmp_path).stat().st_size / 1_048_576
            if result_mb <= max_mb:
                return tmp_path, orig_mb, result_mb
            os.remove(tmp_path)
        except Exception:
            if tmp_path and os.path.exists(tmp_path):
                try: os.remove(tmp_path)
                except: pass

    return path, orig_mb, orig_mb


def _comprimir_pdf(path: str) -> tuple:
    """
    Re-renderiza páginas escaneadas a 110 DPI / JPEG quality 48 (moderado-alto).
    Páginas de texto são copiadas sem perda. Retorna (path_comprimido, mb_original, mb_comprimido).
    Se a compressão não trouxer ganho ou falhar, devolve o path original.
    """
    original_mb = Path(path).stat().st_size / 1_048_576
    tmp_path = None
    try:
        doc = fitz.open(path)
        new_doc = fitz.open()
        for pn, page in enumerate(doc):
            txt = page.get_text().strip()
            alfa = sum(c.isalpha() for c in txt) if txt else 0
            is_text = (
                len(txt) >= 50
                and alfa / max(len(txt), 1) >= 0.4
                and len(txt.split()) >= 30
            )
            if is_text:
                new_doc.insert_pdf(doc, from_page=pn, to_page=pn)
            else:
                mat = fitz.Matrix(110 / 72, 110 / 72)
                pix = page.get_pixmap(matrix=mat)
                img_bytes = pix.tobytes("jpeg", jpg_quality=48)
                rect = page.rect
                new_page = new_doc.new_page(width=rect.width, height=rect.height)
                new_page.insert_image(rect, stream=img_bytes, keep_proportion=False)
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_path = tmp.name
        tmp.close()
        new_doc.save(tmp_path, garbage=4, deflate=True, clean=True)
        new_doc.close()
        doc.close()
        compressed_mb = Path(tmp_path).stat().st_size / 1_048_576
        if compressed_mb >= original_mb * 0.95:
            os.remove(tmp_path)
            return path, original_mb, original_mb
        return tmp_path, original_mb, compressed_mb
    except Exception:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return path, original_mb, original_mb


def _barra_progresso(atual: int, total: int) -> str:
    pct = atual / total if total else 0
    blocos = int(pct * 20)
    barra = "█" * blocos + "░" * (20 - blocos)
    return f"[{barra}] {int(pct*100)}% ({atual}/{total} chunks)"



