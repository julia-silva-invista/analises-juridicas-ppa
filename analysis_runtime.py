# -*- coding: utf-8 -*-
"""Coordenação leve de análises, chunks de PDF e diagnóstico persistente.

O módulo não cria processos nem monitora a máquina continuamente. Ele apenas:
- limita a quatro análises ativas;
- divide seis slots de extração entre elas;
- produz chunks PDF sem rasterizar/recomprimir imagens;
- registra mudanças de etapa e reaproveita extrações por arquivo/faixa de páginas.
"""
from __future__ import annotations

import contextlib
import hashlib
import json
import os
import tempfile
import threading
import time
import uuid
from collections import deque
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Optional


MAX_ACTIVE_ANALYSES = int(os.getenv("MAX_ACTIVE_ANALYSES", "4"))
TOTAL_EXTRACTION_WORKERS = int(os.getenv("TOTAL_EXTRACTION_WORKERS", "6"))
PDF_CHUNK_MAX_MB = float(os.getenv("PDF_CHUNK_MAX_MB", "45"))
_PDF_PREPARATION_SEMAPHORE = threading.Semaphore(
    max(1, int(os.getenv("PDF_PREPARATION_CONCURRENCY", "2")))
)


def _base_dir() -> Path:
    override = os.getenv("ANALYSIS_RUNTIME_DIR")
    if override:
        return Path(override)
    data_dir = Path("/data")
    if data_dir.exists() and data_dir.is_dir():
        return data_dir / "analysis_runtime"
    return Path("resultados") / "analysis_runtime"


BASE_DIR = _base_dir()
_STATUS_DIR = BASE_DIR / "status"
_CHUNK_CACHE_DIR = BASE_DIR / "chunks_v1"
_RETENTION_DAYS = int(os.getenv("ANALYSIS_CACHE_RETENTION_DAYS", "7"))
_cleanup_lock = threading.Lock()
_last_cleanup = 0.0


def _write_atomic(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".{uuid.uuid4().hex}.tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def _maybe_cleanup() -> None:
    """Limpa apenas arquivos próprios antigos, no máximo uma vez por hora."""
    global _last_cleanup
    now = time.time()
    if now - _last_cleanup < 3600 or not _cleanup_lock.acquire(blocking=False):
        return
    try:
        if now - _last_cleanup < 3600:
            return
        cutoff = now - max(1, _RETENTION_DAYS) * 86400
        for root, suffix in ((_STATUS_DIR, ".json"), (_CHUNK_CACHE_DIR, ".txt")):
            if not root.exists():
                continue
            for path in root.rglob(f"*{suffix}"):
                try:
                    if path.is_file() and not path.is_symlink() and path.stat().st_mtime < cutoff:
                        path.unlink()
                except OSError:
                    pass
        _last_cleanup = now
    finally:
        _cleanup_lock.release()


class AnalysisManager:
    """Fila de análises e rateador cooperativo dos slots de extração."""

    def __init__(self, max_active: int = 4, total_workers: int = 6):
        self.max_active = max(1, max_active)
        self.total_workers = max(1, total_workers)
        self._cv = threading.Condition()
        self._waiting: deque[str] = deque()
        self._active: dict[str, str] = {}
        self._running: dict[str, int] = {}
        self._worker_waiters: deque[tuple[str, str]] = deque()

    def create(self, kind: str) -> "AnalysisJob":
        _maybe_cleanup()
        job_id = uuid.uuid4().hex
        with self._cv:
            self._waiting.append(job_id)
            self._cv.notify_all()
        return AnalysisJob(self, job_id, kind)

    def try_activate(self, job_id: str, kind: str) -> tuple[bool, int]:
        with self._cv:
            if job_id in self._active:
                return True, 0
            try:
                position = list(self._waiting).index(job_id) + 1
            except ValueError:
                return False, 0
            if position == 1 and len(self._active) < self.max_active:
                self._waiting.popleft()
                self._active[job_id] = kind
                self._running[job_id] = 0
                self._cv.notify_all()
                return True, 0
            return False, position

    def wait_for_change(self, timeout: float = 2.0) -> None:
        with self._cv:
            self._cv.wait(timeout=timeout)

    def _per_job_cap_locked(self) -> int:
        active = max(1, len(self._active))
        if active == 1:
            return self.total_workers
        if active == 2:
            return min(3, self.total_workers)
        if active == 3:
            return min(2, self.total_workers)
        return 1

    def worker_cap(self) -> int:
        with self._cv:
            return self._per_job_cap_locked()

    def acquire_worker(self, job_id: str) -> str:
        token = uuid.uuid4().hex
        with self._cv:
            self._worker_waiters.append((token, job_id))
            while True:
                if job_id not in self._active:
                    try:
                        self._worker_waiters.remove((token, job_id))
                    except ValueError:
                        pass
                    raise RuntimeError("A análise foi encerrada antes de receber um worker.")
                cap = self._per_job_cap_locked()
                total_running = sum(self._running.values())
                eligible_token = None
                for candidate, candidate_job in self._worker_waiters:
                    if candidate_job in self._active and self._running.get(candidate_job, 0) < cap:
                        eligible_token = candidate
                        break
                if (
                    job_id in self._active
                    and token == eligible_token
                    and self._running.get(job_id, 0) < cap
                    and total_running < self.total_workers
                ):
                    self._worker_waiters.remove((token, job_id))
                    self._running[job_id] = self._running.get(job_id, 0) + 1
                    return token
                self._cv.wait(timeout=1.0)

    def release_worker(self, job_id: str) -> None:
        with self._cv:
            if job_id in self._running:
                self._running[job_id] = max(0, self._running[job_id] - 1)
            self._cv.notify_all()

    def finish(self, job_id: str) -> None:
        with self._cv:
            self._active.pop(job_id, None)
            self._running.pop(job_id, None)
            try:
                self._waiting.remove(job_id)
            except ValueError:
                pass
            self._worker_waiters = deque(
                (token, owner) for token, owner in self._worker_waiters if owner != job_id
            )
            self._cv.notify_all()

    def snapshot(self) -> dict:
        with self._cv:
            return {
                "active": len(self._active),
                "waiting": len(self._waiting),
                "workers_running": sum(self._running.values()),
                "workers_total": self.total_workers,
                "workers_per_analysis": self._per_job_cap_locked() if self._active else self.total_workers,
            }


class AnalysisJob:
    def __init__(self, manager: AnalysisManager, job_id: str, kind: str):
        self.manager = manager
        self.id = job_id
        self.kind = kind
        self._closed = False

    def try_activate(self) -> tuple[bool, int]:
        return self.manager.try_activate(self.id, self.kind)

    def wait_for_change(self, timeout: float = 2.0) -> None:
        self.manager.wait_for_change(timeout)

    @contextlib.contextmanager
    def worker_slot(self):
        self.manager.acquire_worker(self.id)
        try:
            yield
        finally:
            self.manager.release_worker(self.id)

    def close(self) -> None:
        if not self._closed:
            self._closed = True
            self.manager.finish(self.id)


ANALYSIS_MANAGER = AnalysisManager(MAX_ACTIVE_ANALYSES, TOTAL_EXTRACTION_WORKERS)


@contextlib.contextmanager
def pdf_preparation_slot():
    """Evita que quatro PDFs grandes sejam regravados ao mesmo tempo."""
    with _PDF_PREPARATION_SEMAPHORE:
        yield


def environment_status() -> dict:
    snap = ANALYSIS_MANAGER.snapshot()
    demand = snap["active"] + snap["waiting"]
    if demand <= 2:
        state, label = "stable", "Estável"
    elif demand == 3:
        state, label = "operational", "Operacional"
    else:
        state, label = "peak", "Pico de demanda"
    return {**snap, "state": state, "label": label}


def environment_status_json() -> str:
    return json.dumps(environment_status(), ensure_ascii=False)


def queue_message(position: int) -> str:
    suffix = f" Posição atual: {position}." if position else ""
    return (
        "⚠️ Sistema sobrecarregado: já existem 4 análises em andamento. "
        "Sua análise está na fila e em breve será iniciada, assim que houver capacidade."
        + suffix
    )


def record_status(job: Optional[AnalysisJob], stage: str, **details) -> None:
    """Grava poucos bytes somente quando o fluxo muda de etapa."""
    if not job:
        return
    payload = {
        "job_id": job.id,
        "tipo": job.kind,
        "etapa": stage,
        "atualizado_em": time.strftime("%Y-%m-%dT%H:%M:%S"),
        **details,
    }
    try:
        _write_atomic(_STATUS_DIR / f"{job.id}.json", json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception:
        pass


@dataclass(frozen=True)
class PdfChunk:
    path: str
    source_path: str
    start: int
    end: int
    total_pages: int
    temporary: bool

    @property
    def page_start(self) -> int:
        return self.start + 1

    @property
    def page_end(self) -> int:
        return self.end


class OversizedPdfPageError(RuntimeError):
    pass


def iter_pdf_chunks(
    source_path: str,
    max_pages: int = 400,
    max_mb: float = PDF_CHUNK_MAX_MB,
) -> Iterator[PdfChunk]:
    """Gera um chunk por vez, bissectando por tamanho sem rasterizar páginas.

    ``garbage=4`` e ``deflate=True`` apenas limpam/compactam os streams PDF sem alterar
    resolução ou qualidade das imagens. Não há conversão das páginas para JPEG.
    """
    import fitz

    source_path = str(source_path)
    source_size = Path(source_path).stat().st_size
    limit_bytes = int(max_mb * 1024 * 1024)
    doc = fitz.open(source_path)
    total = len(doc)
    if total <= 0:
        doc.close()
        raise ValueError(f"PDF sem páginas: {Path(source_path).name}")

    def produce(start: int, end: int) -> Iterator[PdfChunk]:
        if start == 0 and end == total and total <= max_pages and source_size <= limit_bytes:
            yield PdfChunk(source_path, source_path, start, end, total, False)
            return

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_path = tmp.name
        tmp.close()
        try:
            sub = fitz.open()
            try:
                sub.insert_pdf(doc, from_page=start, to_page=end - 1)
                sub.save(tmp_path, garbage=4, deflate=True)
            finally:
                sub.close()
            size = Path(tmp_path).stat().st_size
            if size <= limit_bytes:
                accepted_path = tmp_path
                tmp_path = ""
                yield PdfChunk(accepted_path, source_path, start, end, total, True)
                return
            if end - start <= 1:
                raise OversizedPdfPageError(
                    f"A página {start + 1} de '{Path(source_path).name}' ainda possui "
                    f"{size / 1_048_576:.1f} MB sozinha (limite seguro: {max_mb:.0f} MB). "
                    "Ela precisa de tratamento manual antes da análise."
                )
            middle = start + (end - start) // 2
            yield from produce(start, middle)
            yield from produce(middle, end)
        finally:
            if tmp_path:
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    try:
        for block_start in range(0, total, max_pages):
            block_end = min(block_start + max_pages, total)
            yield from produce(block_start, block_end)
    finally:
        doc.close()


def cleanup_chunk(chunk: PdfChunk) -> None:
    if chunk.temporary:
        try:
            os.remove(chunk.path)
        except OSError:
            pass


def file_sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as file_obj:
        for block in iter(lambda: file_obj.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def _chunk_cache_path(
    source_hash: str,
    start: int,
    end: int,
    model: str,
    namespace: str,
) -> Path:
    identity = f"{namespace}|{source_hash}|{start}|{end}|{model}|v1"
    key = hashlib.sha256(identity.encode("utf-8")).hexdigest()
    return _CHUNK_CACHE_DIR / source_hash[:2] / f"{key}.txt"


def load_chunk_cache(
    source_hash: str,
    start: int,
    end: int,
    model: str,
    namespace: str,
) -> Optional[str]:
    path = _chunk_cache_path(source_hash, start, end, model, namespace)
    try:
        return path.read_text(encoding="utf-8") if path.exists() else None
    except Exception:
        return None


def save_chunk_cache(
    source_hash: str,
    start: int,
    end: int,
    model: str,
    namespace: str,
    text: str,
) -> None:
    try:
        _write_atomic(_chunk_cache_path(source_hash, start, end, model, namespace), text or "")
    except Exception:
        pass
