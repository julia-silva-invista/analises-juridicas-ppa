import os
import sys
import json
import random
import threading
import time
import tempfile
from contextlib import contextmanager
from pathlib import Path

import fitz

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import analysis_runtime as runtime
from analysis_runtime import AnalysisManager, queue_message
import rj_cache


@contextmanager
def _temporary_files():
    paths = []

    def make(suffix: str = ".pdf") -> Path:
        path = Path(tempfile.mktemp(prefix="analysis_runtime_test_", suffix=suffix))
        paths.append(path)
        return path

    try:
        yield make
    finally:
        for path in paths:
            try:
                path.unlink()
            except FileNotFoundError:
                pass


def _activate(manager, kind="teste"):
    job = manager.create(kind)
    active, position = job.try_activate()
    assert active is True
    assert position == 0
    return job


def test_quinta_analise_fica_na_fila_ate_uma_vaga_abrir():
    manager = AnalysisManager(max_active=4, total_workers=6)
    active_jobs = [_activate(manager) for _ in range(4)]

    fifth = manager.create("rj")
    active, position = fifth.try_activate()
    assert active is False
    assert position == 1
    assert "sobrecarregado" in queue_message(position).lower()
    assert "Posição atual: 1" in queue_message(position)

    active_jobs[0].close()
    active, position = fifth.try_activate()
    assert active is True
    assert position == 0

    for job in active_jobs[1:]:
        job.close()
    fifth.close()


def test_limite_por_analise_muda_com_a_demanda():
    manager = AnalysisManager(max_active=4, total_workers=6)
    jobs = []
    expected = [6, 3, 2, 1]
    for cap in expected:
        jobs.append(_activate(manager))
        assert manager.worker_cap() == cap
    for job in jobs:
        job.close()


def test_seis_workers_globais_e_rebalanceamento_sem_interromper_tarefa():
    manager = AnalysisManager(max_active=4, total_workers=6)
    first = _activate(manager)
    entered = []
    release = threading.Event()

    def use_slot(index):
        with first.worker_slot():
            entered.append(index)
            release.wait(2)

    threads = [threading.Thread(target=use_slot, args=(i,)) for i in range(6)]
    for thread in threads:
        thread.start()

    deadline = time.time() + 2
    while len(entered) < 6 and time.time() < deadline:
        time.sleep(0.01)
    assert len(entered) == 6

    second = _activate(manager)
    assert manager.worker_cap() == 3
    # Os seis slots já iniciados não são cancelados no meio; o novo teto vale
    # para as próximas aquisições.
    assert manager.snapshot()["workers_running"] == 6

    release.set()
    for thread in threads:
        thread.join(timeout=2)
    assert manager.snapshot()["workers_running"] == 0
    first.close()
    second.close()


def test_registro_persistente_e_pequeno_e_guarda_a_faixa_de_paginas():
    manager = AnalysisManager(max_active=4, total_workers=6)
    job = _activate(manager, "rj")
    captured = {}
    original_write = runtime._write_atomic
    try:
        runtime._write_atomic = lambda path, text: captured.update(path=path, text=text)
        runtime.record_status(
            job, "extraindo_chunk", arquivo="P1.PDF", paginas="401-800", chunk=2
        )
        payload = json.loads(captured["text"])
        assert payload["etapa"] == "extraindo_chunk"
        assert payload["arquivo"] == "P1.PDF"
        assert payload["paginas"] == "401-800"
        assert len(captured["text"].encode("utf-8")) < 1024
    finally:
        runtime._write_atomic = original_write
        job.close()


def test_cache_e_independente_por_arquivo_faixa_e_modelo():
    base = runtime._chunk_cache_path("a" * 64, 0, 400, "modelo-a", "rj")
    assert base != runtime._chunk_cache_path("a" * 64, 400, 800, "modelo-a", "rj")
    assert base != runtime._chunk_cache_path("a" * 64, 0, 400, "modelo-b", "rj")
    assert base != runtime._chunk_cache_path("b" * 64, 0, 400, "modelo-a", "rj")
    assert "chunks_v3" in str(base)


def _criar_pdf_pequeno(path: Path, texto: str = "PAGINA ORIGINAL") -> None:
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), texto)
    doc.save(path)
    doc.close()


def _criar_pdf_recursos_compartilhados(path: Path, paginas: int = 6) -> None:
    """Simula o Documento Unificado: todas as páginas enxergam todos os Forms."""
    source = fitz.open()
    samples = bytes(((i * 73 + 19) % 256 for i in range(200 * 200 * 3)))
    pix = fitz.Pixmap(fitz.csRGB, 200, 200, samples, False)
    for indice in range(paginas):
        page = source.new_page(width=300, height=300)
        page.insert_text((20, 20), f"UNIQUE PAGE {indice + 1}")
        page.insert_image(fitz.Rect(20, 30, 280, 290), pixmap=pix)

    unified = fitz.open()
    forms = []
    for indice in range(paginas):
        page = unified.new_page(width=300, height=300)
        forms.append(page.show_pdf_page(page.rect, source, indice))

    resources = unified.get_new_xref()
    xobjects = " ".join(f"/TPL{i} {xref} 0 R" for i, xref in enumerate(forms))
    unified.update_object(resources, f"<< /XObject << {xobjects} >> >>")
    for indice, page in enumerate(unified):
        unified.xref_set_key(page.xref, "Resources", f"{resources} 0 R")
        unified.update_stream(page.get_contents()[0], f"q /TPL{indice} Do Q".encode())

    unified.save(path, garbage=0, deflate=False)
    unified.close()
    source.close()


def _criar_pdf_imagem_incompressivel(path: Path) -> None:
    doc = fitz.open()
    page = doc.new_page(width=600, height=800)
    samples = random.Random(123).randbytes(1000 * 1000 * 3)
    pix = fitz.Pixmap(fitz.csRGB, 1000, 1000, samples, False)
    page.insert_image(page.rect, pixmap=pix)
    doc.save(path, deflate=False)
    doc.close()


def test_pdf_pequeno_permanece_original():
    with _temporary_files() as make:
        source = make("_pequeno.pdf")
        _criar_pdf_pequeno(source)
        chunks = list(runtime.iter_pdf_chunks(str(source)))
        assert len(chunks) == 1
        chunk = chunks[0]
        assert chunk.path == str(source)
        assert chunk.temporary is False
        assert chunk.preparation_mode == "original"
        assert "original preservado" in chunk.preparation_note


def test_recursos_globais_sao_removidos_sem_perder_paginas_ou_texto():
    with _temporary_files() as make:
        source = make("_unificado.pdf")
        _criar_pdf_recursos_compartilhados(source)
        chunks = list(runtime.iter_pdf_chunks(str(source), max_pages=2, max_mb=0.05))
        try:
            assert [(chunk.start, chunk.end) for chunk in chunks] == [(0, 2), (2, 4), (4, 6)]
            assert all(chunk.preparation_mode == "resources_cleaned" for chunk in chunks)
            assert all(Path(chunk.path).stat().st_size <= 0.05 * 1024 * 1024 for chunk in chunks)
            textos = []
            for chunk in chunks:
                with fitz.open(chunk.path) as doc:
                    textos.extend(page.get_text() for page in doc)
            assert [f"UNIQUE PAGE {i}" in textos[i - 1] for i in range(1, 7)] == [True] * 6
        finally:
            temporarios = [Path(chunk.path) for chunk in chunks if chunk.temporary]
            for chunk in chunks:
                runtime.cleanup_chunk(chunk)
            assert all(not path.exists() for path in temporarios)


def test_pagina_realmente_grande_e_rasterizada_ate_caber():
    with _temporary_files() as make:
        source = make("_pagina_grande.pdf")
        _criar_pdf_imagem_incompressivel(source)
        chunks = list(runtime.iter_pdf_chunks(str(source), max_pages=1, max_mb=2))
        try:
            assert len(chunks) == 1
            chunk = chunks[0]
            assert chunk.preparation_mode == "page_rasterized"
            assert chunk.raster_dpi in {180, 144, 108}
            assert Path(chunk.path).stat().st_size <= 2 * 1024 * 1024
            assert "rasterizada" in chunk.preparation_note
        finally:
            for chunk in chunks:
                runtime.cleanup_chunk(chunk)


def test_temporarios_sao_removidos_quando_pagina_nao_cabe():
    with _temporary_files() as make:
        source = make("_pagina_impossivel.pdf")
        _criar_pdf_imagem_incompressivel(source)
        created = []
        original_named_temporary_file = runtime.tempfile.NamedTemporaryFile

        def tracked_named_temporary_file(*args, **kwargs):
            result = original_named_temporary_file(*args, **kwargs)
            created.append(Path(result.name))
            return result

        runtime.tempfile.NamedTemporaryFile = tracked_named_temporary_file
        try:
            try:
                list(runtime.iter_pdf_chunks(str(source), max_pages=1, max_mb=0.5))
                raise AssertionError("Era esperado OversizedPdfPageError")
            except runtime.OversizedPdfPageError:
                pass
            assert created
            assert all(not path.exists() for path in created)
        finally:
            runtime.tempfile.NamedTemporaryFile = original_named_temporary_file


def test_detecta_apenas_excesso_de_tokens_de_entrada():
    assert runtime.input_token_limit_exceeded(
        RuntimeError("400 INVALID_ARGUMENT: The input token count exceeds the maximum number of tokens allowed 1048576")
    )
    assert not runtime.input_token_limit_exceeded(
        RuntimeError("400 INVALID_ARGUMENT: Request contains an invalid argument")
    )


def test_subdivisao_por_tokens_preserva_faixas_e_limpa_temporarios():
    with _temporary_files() as make:
        source = make("_cinco_paginas.pdf")
        doc = fitz.open()
        for indice in range(5):
            page = doc.new_page()
            page.insert_text((72, 72), f"PAGINA {indice + 1}")
        doc.save(source)
        doc.close()

        chunks = runtime.split_pdf_chunk_for_token_limit(
            str(source), absolute_start=400, total_pages=1000, source_path="P1.PDF"
        )
        temporarios = [Path(chunk.path) for chunk in chunks]
        try:
            assert [(chunk.start, chunk.end) for chunk in chunks] == [(400, 403), (403, 405)]
            assert all(chunk.total_pages == 1000 for chunk in chunks)
            assert all(chunk.preparation_mode == "token_split" for chunk in chunks)
            assert all(chunk.temporary for chunk in chunks)
        finally:
            for chunk in chunks:
                runtime.cleanup_chunk(chunk)
        assert all(not path.exists() for path in temporarios)


def test_versao_do_pipeline_invalida_job_id_de_rj():
    with _temporary_files() as make:
        source = make("_rj.pdf")
        _criar_pdf_pequeno(source)
        original = rj_cache.ANALYSIS_PIPELINE_VERSION
        try:
            first = rj_cache.calcular_job_id(
                [str(source)], [], "", False, "extracao", "relatorio"
            )
            rj_cache.ANALYSIS_PIPELINE_VERSION = original + "-outro"
            second = rj_cache.calcular_job_id(
                [str(source)], [], "", False, "extracao", "relatorio"
            )
            assert first != second
        finally:
            rj_cache.ANALYSIS_PIPELINE_VERSION = original


if __name__ == "__main__":
    test_quinta_analise_fica_na_fila_ate_uma_vaga_abrir()
    test_limite_por_analise_muda_com_a_demanda()
    test_seis_workers_globais_e_rebalanceamento_sem_interromper_tarefa()
    test_registro_persistente_e_pequeno_e_guarda_a_faixa_de_paginas()
    test_cache_e_independente_por_arquivo_faixa_e_modelo()
    test_pdf_pequeno_permanece_original()
    test_recursos_globais_sao_removidos_sem_perder_paginas_ou_texto()
    test_pagina_realmente_grande_e_rasterizada_ate_caber()
    test_temporarios_sao_removidos_quando_pagina_nao_cabe()
    test_detecta_apenas_excesso_de_tokens_de_entrada()
    test_subdivisao_por_tokens_preserva_faixas_e_limpa_temporarios()
    test_versao_do_pipeline_invalida_job_id_de_rj()
    print("12 testes de coordenação/PDF/cache: OK")
