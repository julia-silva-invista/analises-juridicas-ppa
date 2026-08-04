import os
import sys
import json
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import analysis_runtime as runtime
from analysis_runtime import AnalysisManager, queue_message


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


if __name__ == "__main__":
    test_quinta_analise_fica_na_fila_ate_uma_vaga_abrir()
    test_limite_por_analise_muda_com_a_demanda()
    test_seis_workers_globais_e_rebalanceamento_sem_interromper_tarefa()
    test_registro_persistente_e_pequeno_e_guarda_a_faixa_de_paginas()
    test_cache_e_independente_por_arquivo_faixa_e_modelo()
    print("5 testes de coordenação/cache: OK")
