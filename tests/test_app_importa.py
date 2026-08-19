"""O `app.py` tem de montar a interface inteira na versão de Gradio que o Space usa.

Por que existe: os bugs mais recorrentes do repositório são os que só apareciam depois do
deploy ("exportar imagem da timeline dispara no Gradio 5 do Space", "compatibiliza
google-genai com Gradio do Space"). Todos eram erro na *construção* da interface, que este
teste pega antes de publicar — basta rodar a suíte no ambiente com a versão do Space.

O `app.py` chama `demo.launch(...)` no nível do módulo, então o teste substitui `launch` por
uma função inerte antes de importar: monta tudo, sem subir servidor.
"""
from __future__ import annotations

import sys
from pathlib import Path

import gradio as gr

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def test_app_monta_a_interface_sem_subir_servidor(monkeypatch):
    lancamentos = []
    monkeypatch.setattr(
        gr.Blocks, "launch", lambda self, *a, **kw: lancamentos.append(kw) or self
    )
    sys.modules.pop("app", None)

    import app  # noqa: F401  — o import é o próprio teste

    assert lancamentos, "app.py deveria ter chamado demo.launch()"
    assert isinstance(app.demo, gr.Blocks)


def test_todas_as_abas_esperadas_existem(monkeypatch):
    monkeypatch.setattr(gr.Blocks, "launch", lambda self, *a, **kw: self)
    sys.modules.pop("app", None)

    import app

    rotulos = {
        getattr(bloco, "label", None)
        for bloco in app.demo.blocks.values()
        if type(bloco).__name__ == "Tab"
    }
    esperadas = {
        "Processos",
        "Recuperação Judicial",
        "Matrículas",
        "Timeline Societária",
        "Coleta de Informações",
    }
    faltando = esperadas - rotulos
    assert not faltando, f"abas ausentes: {sorted(faltando)}"
