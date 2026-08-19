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


def test_toda_caixa_de_upload_tem_o_acabamento_padrao(monkeypatch):
    """Upload de entrada usa `inv-upload-box`, que dá a borda e o raio do card de instruções.

    Sem isso a caixa nasce com o canto quadrado do Gradio, e a aba nova sai diferente das
    outras — o tipo de detalhe que costuma virar `fix:` no commit seguinte.
    """
    monkeypatch.setattr(gr.Blocks, "launch", lambda self, *a, **kw: self)
    sys.modules.pop("app", None)

    import app

    entradas_sem_classe = []
    for bloco in app.demo.blocks.values():
        if type(bloco).__name__ != "File":
            continue
        classes = list(getattr(bloco, "elem_classes", None) or [])
        # saídas de download têm acabamento próprio (compact-file-output)
        if "compact-file-output" in classes:
            continue
        if "inv-upload-box" not in classes:
            entradas_sem_classe.append(getattr(bloco, "label", "<sem rótulo>"))

    assert not entradas_sem_classe, (
        "caixas de upload sem 'inv-upload-box': " + ", ".join(map(str, entradas_sem_classe))
    )


def test_todo_modulo_tem_caixa_de_log_com_acabamento_padrao(monkeypatch):
    """As caixas de progresso usam `log-area`, que dá o mesmo acabamento do card.

    São cinco, uma por módulo. Rótulos variam ("Log de execução", "Status", "Log"), então a
    classe é o que garante o desenho uniforme.
    """
    monkeypatch.setattr(gr.Blocks, "launch", lambda self, *a, **kw: self)
    sys.modules.pop("app", None)

    import app

    com_classe = [
        bloco
        for bloco in app.demo.blocks.values()
        if type(bloco).__name__ == "Textbox"
        and "log-area" in list(getattr(bloco, "elem_classes", None) or [])
    ]
    assert len(com_classe) == 5, f"esperava 5 caixas de log, encontrei {len(com_classe)}"
