"""A caixa de perguntas das três abas (Processos, RJ, Matrículas) é a mesma peça.

O botão "Perguntar" nascia grudado no topo da coluna, alinhado com o rótulo em vez da
caixa de texto. A correção é de CSS, então o que se testa é a ligação entre o layout e a
regra: as três abas marcadas com as mesmas classes, e as classes existindo no CSS. Abas
divergindo entre si é exatamente o que regride sem ninguém notar.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import design  # noqa: E402

APP = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")


def testar_as_tres_abas_usam_a_mesma_marcacao_de_pergunta():
    botoes = re.findall(r"(\w+)_perguntar_btn = gr\.Button", APP)
    assert sorted(botoes) == ["mat", "proc", "rj"], botoes
    assert APP.count('with gr.Row(elem_classes=["qa-ask-row"]):') == len(botoes)
    assert APP.count('with gr.Column(scale=1, elem_classes=["qa-ask-col"]):') == len(botoes)


def testar_css_centraliza_o_botao_na_caixa_de_texto():
    """Centralizar na coluna inteira alinharia o botão com o rótulo + a caixa; o recuo do
    rótulo tira o rótulo da conta e sobra a altura da caixa."""
    regra = design.CSS.split(".qa-ask-col {", 1)[1].split("}", 1)[0]
    assert "align-items: center" in regra
    assert "padding-top" in regra, "sem o recuo do rótulo, centraliza na coluna errada"
