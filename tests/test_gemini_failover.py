"""Testes unitários do roteamento entre credenciais Gemini, sem rede."""
from __future__ import annotations

import sys
import types as pytypes
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

# O runtime local leve não possui PyMuPDF/google-genai; o failover não depende
# deles, então stubs mínimos bastam para importar utils.
sys.modules.setdefault("fitz", pytypes.ModuleType("fitz"))

google_mod = pytypes.ModuleType("google")
genai_mod = pytypes.ModuleType("google.genai")
genai_types_mod = pytypes.ModuleType("google.genai.types")
genai_types_mod.HttpOptions = lambda **kwargs: kwargs
genai_mod.types = genai_types_mod
genai_mod.Client = object
google_mod.genai = genai_mod
sys.modules.setdefault("google", google_mod)
sys.modules.setdefault("google.genai", genai_mod)
sys.modules.setdefault("google.genai.types", genai_types_mod)

from utils import _erro_gemini_permite_failover, _executar_com_failover_gemini  # noqa: E402


def testar_401_troca_para_proxima_credencial():
    clients = [object(), object(), object()]
    ordem = []
    avisos = []

    def chamada(_client, indice):
        ordem.append(indice)
        if indice == 0:
            raise RuntimeError("401 UNAUTHENTICATED ACCESS_TOKEN_TYPE_UNSUPPORTED")
        return f"ok-{indice}"

    resultado = _executar_com_failover_gemini(
        clients,
        chamada,
        ao_falhar=lambda atual, proximo, _exc: avisos.append((atual, proximo)),
    )

    assert resultado == "ok-1"
    assert ordem == [0, 1]
    assert avisos == [(0, 1)]


def testar_erro_de_codigo_nao_e_mascarado():
    ordem = []

    def chamada(_client, indice):
        ordem.append(indice)
        raise ValueError("schema local incompatível")

    try:
        _executar_com_failover_gemini([object(), object()], chamada)
    except ValueError as exc:
        assert "schema local" in str(exc)
    else:
        raise AssertionError("O erro local deveria ter sido propagado.")

    assert ordem == [0]


def testar_classificacao_de_erros_e_indice_inicial():
    assert _erro_gemini_permite_failover(RuntimeError("429 RESOURCE_EXHAUSTED"))
    assert _erro_gemini_permite_failover(RuntimeError("404 model is no longer available"))
    assert not _erro_gemini_permite_failover(ValueError("JSON inválido produzido localmente"))

    ordem = []
    resultado = _executar_com_failover_gemini(
        [object(), object(), object()],
        lambda _client, indice: ordem.append(indice) or indice,
        indice_inicial=2,
    )
    assert resultado == 2
    assert ordem == [2]


if __name__ == "__main__":
    testar_401_troca_para_proxima_credencial()
    testar_erro_de_codigo_nao_e_mascarado()
    testar_classificacao_de_erros_e_indice_inicial()
    print("OK — failover Gemini")
