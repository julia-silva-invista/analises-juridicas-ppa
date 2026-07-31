# -*- coding: utf-8 -*-
"""
Checkpointing em disco para analises de RJ muito grandes (dezenas de milhares
de paginas). Permite retomar uma analise interrompida (restart do Space, OOM,
queda de rede) sem reprocessar chunks/lotes ja concluidos.
"""
import os
import json
import time
import shutil
import hashlib
from pathlib import Path
from typing import Optional

BASE_DIR_RJ_JOBS = Path("resultados") / "rj_jobs"
RETENCAO_DIAS_RJ = int(os.getenv("RJ_JOB_RETENCAO_DIAS", "7"))
RETENCAO_DIAS_RJ_CONCLUIDO = int(os.getenv("RJ_JOB_RETENCAO_DIAS_CONCLUIDO", "2"))


def calcular_job_id(pdf_paths, instrucoes: str, versao_resumida: bool, usar_gemini_pro: bool) -> str:
    """Hash estavel: conteudo dos PDFs (ordenados) + parametros que mudam o resultado.
    Usa bytes dos arquivos (nao nome/caminho, que muda a cada upload no Gradio)
    para sobreviver a reenvio do mesmo PDF apos reinicio do processo."""
    h = hashlib.sha256()
    for p in sorted(pdf_paths):
        h.update(Path(p).name.encode("utf-8", "ignore"))
        with open(p, "rb") as f:
            while True:
                chunk = f.read(1024 * 1024)
                if not chunk:
                    break
                h.update(chunk)
    h.update((instrucoes or "").strip().encode("utf-8", "ignore"))
    h.update(b"resumida" if versao_resumida else b"completa")
    h.update(b"pro" if usar_gemini_pro else b"flash")
    return h.hexdigest()[:20]


def caminho_job(job_id: str) -> Path:
    d = BASE_DIR_RJ_JOBS / job_id
    (d / "chunks").mkdir(parents=True, exist_ok=True)
    (d / "lotes").mkdir(parents=True, exist_ok=True)
    (d / "pdf_chunks").mkdir(parents=True, exist_ok=True)
    return d


def _escrever_atomico(caminho: Path, conteudo: str) -> None:
    tmp = caminho.with_suffix(caminho.suffix + ".tmp")
    tmp.write_text(conteudo, encoding="utf-8")
    os.replace(tmp, caminho)


def carregar_manifest(job_id: str) -> Optional[dict]:
    caminho = caminho_job(job_id) / "manifest.json"
    if not caminho.exists():
        return None
    try:
        return json.loads(caminho.read_text(encoding="utf-8"))
    except Exception:
        return None


def salvar_manifest(job_id: str, manifest: dict) -> None:
    manifest["atualizado_em"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    caminho = caminho_job(job_id) / "manifest.json"
    _escrever_atomico(caminho, json.dumps(manifest, ensure_ascii=False, indent=2))


def novo_manifest(job_id: str, arquivos: list, n_chunks: int, chunk_max_pages: int) -> dict:
    return {
        "job_id": job_id,
        "criado_em": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "atualizado_em": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "arquivos": arquivos,
        "n_chunks": n_chunks,
        "chunk_max_pages": chunk_max_pages,
        "arquivos_status": {},   # nome do arquivo -> "ok" (dividido/comprimido) — fase de divisao
        "chunks_index": [],      # [{"arquivo":..., "offset":..., "total":..., "chunk_pdf": "pdf_chunks/xx.pdf"}, ...]
        "chunks_status": {},     # str(indice) -> "ok"/"erro"/"erro_definitivo" — fase de extracao
        "lotes_status": {},
        "secao_a_ok": False,
        "concluido": False,
    }


def salvar_pdf_chunk(job_id: str, nome_rel: str, src_path: str) -> str:
    """Copia um PDF-chunk (resultado de _rj_dividir_pdf) para dentro do diretorio
    do job, para sobreviver a reinicio do container (diferente de tempfile no /tmp
    do SO, que pode ser perdido se o Space reiniciar)."""
    dest = caminho_job(job_id) / "pdf_chunks" / nome_rel
    if os.path.abspath(src_path) != os.path.abspath(dest):
        shutil.copyfile(src_path, dest)
    return str(dest)


def caminho_pdf_chunk(job_id: str, nome_rel: str) -> str:
    return str(caminho_job(job_id) / "pdf_chunks" / nome_rel)


def salvar_chunk(job_id: str, idx: int, texto: str) -> None:
    caminho = caminho_job(job_id) / "chunks" / f"chunk_{idx:04d}.txt"
    _escrever_atomico(caminho, texto or "")


def carregar_chunk(job_id: str, idx: int) -> str:
    caminho = caminho_job(job_id) / "chunks" / f"chunk_{idx:04d}.txt"
    return caminho.read_text(encoding="utf-8") if caminho.exists() else ""


def salvar_lote(job_id: str, idx: int, texto: str) -> None:
    caminho = caminho_job(job_id) / "lotes" / f"lote_{idx:04d}.txt"
    _escrever_atomico(caminho, texto or "")


def carregar_lote(job_id: str, idx: int) -> str:
    caminho = caminho_job(job_id) / "lotes" / f"lote_{idx:04d}.txt"
    return caminho.read_text(encoding="utf-8") if caminho.exists() else ""


def limpar_jobs_antigos(base_dir: Path = BASE_DIR_RJ_JOBS) -> None:
    """Remove diretorios de job cujo manifest esteja mais velho que a retencao
    configurada (jobs concluidos expiram mais rapido que incompletos, ja que
    nao ha motivo pratico para retomar um job ja finalizado). Nao bloqueante:
    erro aqui nunca deve interromper a analise principal."""
    if not base_dir.exists():
        return
    agora = time.time()
    for job_dir in base_dir.iterdir():
        try:
            manifest_path = job_dir / "manifest.json"
            if not manifest_path.exists():
                continue
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            dias = RETENCAO_DIAS_RJ_CONCLUIDO if manifest.get("concluido") else RETENCAO_DIAS_RJ
            limite = agora - dias * 86400
            if manifest_path.stat().st_mtime < limite:
                shutil.rmtree(job_dir, ignore_errors=True)
        except Exception:
            pass
