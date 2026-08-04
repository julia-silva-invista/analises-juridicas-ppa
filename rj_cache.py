# -*- coding: utf-8 -*-
"""
Cache em disco dos trechos (chunks) já extraídos numa análise de RJ, para que uma falha no meio
do processamento não jogue fora o que já foi obtido do Gemini (a parte cara/lenta).

Usa um Storage Bucket montado em /data (persistente entre restarts do Space) quando disponível;
caso contrário cai para resultados/ local (comportamento também usado fora do Space).
"""
import hashlib
import json
import os
import shutil
import time
from pathlib import Path
from typing import Optional

from analysis_runtime import ANALYSIS_PIPELINE_VERSION


def _base_dir() -> Path:
    override = os.getenv("RJ_CACHE_DIR")
    if override:
        return Path(override)
    data_dir = Path("/data")
    if data_dir.exists() and data_dir.is_dir():
        return data_dir / "rj_cache"
    return Path("resultados") / "rj_cache"


BASE_DIR = _base_dir()
RETENCAO_DIAS = int(os.getenv("RJ_CACHE_RETENCAO_DIAS", "7"))
RETENCAO_DIAS_CONCLUIDO = int(os.getenv("RJ_CACHE_RETENCAO_DIAS_CONCLUIDO", "2"))


def calcular_job_id(pdf_paths_principal: list, pdf_paths_relacionados: list, instrucoes: str,
                     versao_resumida: bool, model_extracao: str, model_relatorio: str) -> str:
    """Hash estável: conteúdo dos PDFs (principal + relacionados, ordenados) + parâmetros que
    mudam o resultado. Usa bytes dos arquivos (não nome/caminho, que muda a cada upload no
    Gradio) para sobreviver a reenvio do mesmo PDF."""
    h = hashlib.sha256()
    for grupo, paths in (("principal", pdf_paths_principal or []), ("relacionados", pdf_paths_relacionados or [])):
        h.update(grupo.encode("ascii"))
        for p in sorted(paths, key=lambda x: Path(x).name):
            h.update(Path(p).name.encode("utf-8", "ignore"))
            with open(p, "rb") as f:
                while True:
                    chunk = f.read(1024 * 1024)
                    if not chunk:
                        break
                    h.update(chunk)
    h.update((instrucoes or "").strip().encode("utf-8", "ignore"))
    h.update(b"resumida" if versao_resumida else b"completa")
    h.update((model_extracao or "").encode("utf-8", "ignore"))
    h.update((model_relatorio or "").encode("utf-8", "ignore"))
    h.update(ANALYSIS_PIPELINE_VERSION.encode("ascii"))
    return h.hexdigest()[:20]


def caminho_job(job_id: str) -> Path:
    d = BASE_DIR / job_id
    (d / "chunks_principal").mkdir(parents=True, exist_ok=True)
    (d / "chunks_relacionados").mkdir(parents=True, exist_ok=True)
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
    _escrever_atomico(caminho_job(job_id) / "manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))


def novo_manifest(job_id: str) -> dict:
    return {
        "job_id": job_id,
        "criado_em": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "chunks_status_principal": {},     # str(indice) -> "ok"
        "chunks_status_relacionados": {},  # str(indice) -> "ok"
        "concluido": False,
    }


def salvar_chunk(job_id: str, grupo: str, idx: int, texto: str) -> None:
    caminho = caminho_job(job_id) / f"chunks_{grupo}" / f"chunk_{idx:04d}.txt"
    _escrever_atomico(caminho, texto or "")


def carregar_chunk(job_id: str, grupo: str, idx: int) -> str:
    caminho = caminho_job(job_id) / f"chunks_{grupo}" / f"chunk_{idx:04d}.txt"
    return caminho.read_text(encoding="utf-8") if caminho.exists() else ""


def salvar_secao(job_id: str, nome: str, texto: str) -> None:
    _escrever_atomico(caminho_job(job_id) / f"{nome}.txt", texto or "")


def carregar_secao(job_id: str, nome: str) -> Optional[str]:
    caminho = caminho_job(job_id) / f"{nome}.txt"
    return caminho.read_text(encoding="utf-8") if caminho.exists() else None


def limpar_jobs_antigos(base_dir: Path = BASE_DIR) -> None:
    """Remove diretórios de job cujo manifest esteja mais velho que a retenção configurada.
    Não bloqueante: erro aqui nunca deve interromper a análise principal."""
    if not base_dir.exists():
        return
    agora = time.time()
    for job_dir in base_dir.iterdir():
        try:
            manifest_path = job_dir / "manifest.json"
            if not manifest_path.exists():
                continue
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            dias = RETENCAO_DIAS_CONCLUIDO if manifest.get("concluido") else RETENCAO_DIAS
            limite = agora - dias * 86400
            if manifest_path.stat().st_mtime < limite:
                shutil.rmtree(job_dir, ignore_errors=True)
        except Exception:
            pass
