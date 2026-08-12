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

import utils  # noqa: E402
from utils import (  # noqa: E402
    _codigo_http_gemini,
    _detalhe_erro_gemini,
    _erro_gemini_permite_failover,
    _executar_com_failover_gemini,
)


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


def testar_codigo_http_vem_do_status_e_nao_de_substring():
    """Casar "404" em qualquer lugar do texto dá falso positivo: IDs de arquivo,
    contagens de token e URLs entram na mensagem. Um 401 já foi anunciado como
    "modelo não disponível" por causa disso."""
    assert _codigo_http_gemini(RuntimeError("429 RESOURCE_EXHAUSTED. {...}")) == 429

    class _ErroDaApi(Exception):
        code = 401
        message = "Request had invalid authentication credentials."

    assert _codigo_http_gemini(_ErroDaApi()) == 401
    # O ".code" manda mesmo quando o texto tem outros números soltos.
    assert _codigo_http_gemini(RuntimeError("upload de files/404abc com 4040 tokens")) is None
    assert not _erro_gemini_permite_failover(RuntimeError("upload de files/404abc falhou"))


def testar_detalhe_do_erro_e_uma_linha_legivel():
    class _ErroDaApi(Exception):
        code = 404
        message = "models/gemini-x is not found\n  for API version v1beta."

    assert _detalhe_erro_gemini(_ErroDaApi()) == "models/gemini-x is not found for API version v1beta."


if __name__ == "__main__":
    testar_401_troca_para_proxima_credencial()
    testar_erro_de_codigo_nao_e_mascarado()
    testar_classificacao_de_erros_e_indice_inicial()
    testar_codigo_http_vem_do_status_e_nao_de_substring()
    testar_detalhe_do_erro_e_uma_linha_legivel()
    print("OK — failover Gemini")


# ── Falha de rede é transitória, não é conteúdo ──────────────────────────────
# Um piscar de DNS no container derrubava a análise inteira: o chunk falhava, o guarda
# de cobertura bloqueava o relatório parcial (corretamente) e minutos de processamento
# iam embora. O _retry via "[Errno -3] Temporary failure in name resolution" e não
# reconhecia como retentável.

def testar_dns_agora_e_retentado(monkeypatch):
    monkeypatch.setattr(utils.time, "sleep", lambda *_a, **_k: None)
    tentativas = []

    def _falha_dns():
        tentativas.append(1)
        raise RuntimeError("[Errno -3] Temporary failure in name resolution")

    try:
        utils._retry(_falha_dns, tentativas=4, espera_base=1)
    except RuntimeError:
        pass
    assert len(tentativas) == 4, "erro de DNS tem que esgotar as tentativas"


def testar_erro_de_transporte_e_reconhecido_pelo_tipo():
    """str(httpx.ConnectError(...)) não contém o nome da classe — casar só por texto
    deixava passar a família inteira de erros de transporte."""
    class ConnectError(Exception):
        pass

    class ReadTimeout(Exception):
        pass

    assert utils._erro_de_rede(ConnectError("qualquer mensagem"))
    assert utils._erro_de_rede(ReadTimeout(""))
    assert utils._erro_de_rede(RuntimeError("getaddrinfo failed"))
    assert utils._erro_de_rede(RuntimeError("Connection refused"))


def testar_erro_de_conteudo_continua_sem_retry(monkeypatch):
    monkeypatch.setattr(utils.time, "sleep", lambda *_a, **_k: None)
    tentativas = []

    def _schema_ruim():
        tentativas.append(1)
        raise ValueError("schema local incompatível")

    try:
        utils._retry(_schema_ruim, tentativas=4, espera_base=1)
    except ValueError:
        pass
    assert len(tentativas) == 1, "erro de programação não pode ser mascarado por retry"
    assert not utils._erro_de_rede(ValueError("JSON inválido produzido localmente"))


def testar_teto_de_gasto_continua_sem_retry(monkeypatch):
    """Não virou erro de rede por engano — insistir não devolve cota."""
    monkeypatch.setattr(utils.time, "sleep", lambda *_a, **_k: None)
    tentativas = []

    def _teto():
        tentativas.append(1)
        raise RuntimeError("429 RESOURCE_EXHAUSTED: exceeded its monthly spending cap")

    try:
        utils._retry(_teto, tentativas=4, espera_base=1)
    except RuntimeError:
        pass
    assert len(tentativas) == 1
