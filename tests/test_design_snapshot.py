"""Trava de layout: o que o Gradio recebe de `design` não pode mudar sem intenção.

Estas impressões digitais foram tiradas do estado em produção (commit 935bec1), antes de o
logo e o CSS saírem de dentro do `design.py` para `assets/`. A extração é puramente de
armazenamento: se qualquer hash abaixo mudar, o layout mudou — e a mudança tem de ser
deliberada, com o hash atualizado no mesmo commit.
"""
from __future__ import annotations

import hashlib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import design  # noqa: E402

# (tamanho em caracteres, sha256 do texto em UTF-8)
# CSS atualizado deliberadamente: zoom do app de 0.8 para 0.64 (20% menor), com width e
# max-width compensando o zoom para o app seguir preenchendo a largura, como faz o zoom do
# navegador. A trava acusou a mudança nas duas vezes, como deveria.
GOLDEN = {
    "CSS": (66233, "906934a7746bfe8856d03ec654b469291e197bf4c436b7f4b3444e7ca6c49d34"),
    "HEADER_HTML": (479890, "99404ab984ea8d21069d3689de09c9d83ed2757affa9e71680dba93782eabf7b"),
    "FOOTER_HTML": (81, "81f1b568511f1ef005d531eadc1f3318c126037e9ad17654a171cee0630d8eaa"),
}


def _impressao(texto: str) -> tuple[int, str]:
    return len(texto), hashlib.sha256(texto.encode("utf-8")).hexdigest()


def test_saida_de_design_permanece_identica():
    divergentes = []
    for nome, esperado in GOLDEN.items():
        obtido = _impressao(getattr(design, nome))
        if obtido != esperado:
            divergentes.append(
                f"{nome}: esperado {esperado[0]} chars / {esperado[1][:12]}…, "
                f"obtido {obtido[0]} chars / {obtido[1][:12]}…"
            )
    assert not divergentes, "layout alterado:\n" + "\n".join(divergentes)


def test_logo_do_cabecalho_continua_embutido_no_html():
    """O cabeçalho serve a imagem como data URI — o Space não depende de rota de arquivo."""
    assert 'src="data:image/png;base64,' in design.HEADER_HTML
    assert "inv-logo-img" in design.HEADER_HTML
