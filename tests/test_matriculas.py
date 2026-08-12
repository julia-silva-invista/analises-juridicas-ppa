"""Testes locais das regras determinísticas do fluxo de Matrículas."""
from __future__ import annotations

import inspect
import sys
import types as pytypes
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


# O runtime leve de testes não instala google-genai. As regras testadas aqui não
# fazem chamadas externas, portanto basta expor os tipos usados na importação.
google_mod = pytypes.ModuleType("google")
genai_mod = pytypes.ModuleType("google.genai")
genai_types_mod = pytypes.ModuleType("google.genai.types")
genai_types_mod.GenerateContentConfig = lambda **kwargs: kwargs
genai_mod.types = genai_types_mod
google_mod.genai = genai_mod
sys.modules.setdefault("google", google_mod)
sys.modules.setdefault("google.genai", genai_mod)
sys.modules.setdefault("google.genai.types", genai_types_mod)

utils_stub = pytypes.ModuleType("utils")
utils_stub._get_gemini_clients = lambda: []
utils_stub._erro_gemini_permite_failover = lambda exc: False
utils_stub._paginas_digitalizadas_pdf = lambda path: []
utils_stub._retry = lambda fn, **kwargs: fn()
utils_stub._responder_pergunta_generica = lambda *args, **kwargs: ""
utils_stub.GEMINI_MODEL_EXTRACAO = "gemini-extracao-teste"
utils_stub.GEMINI_MODEL_OCR = "gemini-ocr-teste"
utils_stub.GEMINI_MODEL_RELATORIO = "gemini-relatorio-teste"
utils_stub.GEMINI_MODEL_QA = "gemini-qa-teste"
sys.modules["utils"] = utils_stub

import matriculas as mat  # noqa: E402


def testar_modelos_por_etapa():
    fonte_extracao = inspect.getsource(mat._mat_analisar_pdf)
    fonte_consolidacao = inspect.getsource(mat._mat_consolidar_extracao)
    assert "GEMINI_MODEL_EXTRACAO" in fonte_extracao
    assert "GEMINI_MODEL_OCR" in fonte_extracao
    assert "paginas_digitalizadas" in fonte_extracao
    assert "thinking_budget" not in fonte_extracao
    assert "GEMINI_MODEL_RELATORIO" in fonte_consolidacao
    assert "GEMINI_MODEL_QA" in inspect.getsource(mat.mat_responder)
    assert mat.COLUNAS_RENAME_MAT["fracao_ideal"] == "Fração Ideal"


def testar_limite_de_workers():
    fonte = inspect.getsource(mat.mat_gerar_excel)
    assert 'min(n, len(clients), MATRICULAS_MAX_WORKERS)' in fonte
    assert mat.MATRICULAS_MAX_WORKERS == 3


def testar_regras_do_prompt():
    prompt = mat.PROMPT_MATRICULA
    assert "ULTIMA transmissao" in prompt
    assert "arrematacao" in prompt
    assert "dacao em pagamento" in prompt
    assert "exatamente um campo textual" in prompt
    assert "deu origem" in prompt
    assert "fracao_ideal" in prompt
    assert "moeda historica" in prompt
    assert "página absoluta" in prompt
    assert "(R.5 | fl. 8)" in prompt
    assert "PARTE, chunk, texto bruto" in prompt
    assert "**" not in prompt


def testar_formatacao_deterministica():
    resultado = mat._mat_pos_processar({
        "matricula": "21.101 - ENCERRADA, DEU ORIGEM À MATRÍCULA 73.923",
        "comarca": "1º CRI DE CURITIBA/PR",
        "proprietario_atual": (
            "**EMPRESA XPTO S/A CNPJ 12345678000190 - 60%; "
            "JOAO DA SILVA (CPF 123.456.789-00) - 40%**"
        ),
        "fracao_ideal": "EMPRESA XPTO: 60%; JOAO DA SILVA: 40%",
        "descricao_imovel": "IMOVEL URBANO",
        "transmissoes_averbadas_registradas": "R.4: COMPRA E VENDA PARA EMPRESA XPTO CNPJ 12345678000190.",
        "onus_vigentes_registrados_averbados": "AV.5: HIPOTECA EM FAVOR DO BANCO CNPJ 11222333000144.",
        "observacoes": "SEM OBSERVACOES RELEVANTES.",
        "onus_cancelados": "SEM ONUS CANCELADOS IDENTIFICADOS.",
        "grau_confianca": "ALTO",
    })
    assert "**" not in resultado["proprietario_atual"]
    assert "(CNPJ 12.345.678/0001-90)" in resultado["proprietario_atual"], resultado["proprietario_atual"]
    assert "(CPF 123.456.789-00)" in resultado["proprietario_atual"], resultado["proprietario_atual"]
    assert "((CPF" not in resultado["proprietario_atual"]
    assert resultado["matricula"].startswith("21.101 - Encerrada")


def testar_moeda_historica_nao_calculada():
    onus = [
        {
            "codigo": "R.2",
            "tipo": "Hipoteca",
            "valor_principal": 5_000_000,
            "moeda": "Cr$",
            "data_celebracao": "10/01/1990",
            "cancelado": False,
        },
        {
            "codigo": "R.8",
            "tipo": "Hipoteca",
            "valor_principal": 100_000,
            "moeda": "BRL",
            "data_celebracao": "01/01/2024",
            "cancelado": False,
        },
    ]
    calculo = mat._mat_calcular_valor_onus(
        onus,
        onus_vigentes_texto="R.2: Hipoteca em Cr$, vigente.\n\nR.8: Hipoteca em reais, vigente.",
        data_referencia=date(2024, 1, 1),
    )
    assert "R.8" in calculo
    assert "R.2" not in calculo


if __name__ == "__main__":
    testar_modelos_por_etapa()
    testar_limite_de_workers()
    testar_regras_do_prompt()
    testar_formatacao_deterministica()
    testar_moeda_historica_nao_calculada()
    print("5 testes de Matrículas: OK")


# ── Partes da execução: nome passou a valer por si ──────────────────────────
# Antes, o parse devolvia só um set de documentos e o nome digitado era descartado:
# quem informasse apenas o nome não recebia alerta nenhum, sem aviso.

def _transmissao(**campos) -> dict:
    base = {"de_nome": "", "para_nome": "", "de_doc": "", "para_doc": "", "data": ""}
    return {"transmissoes_estruturadas": [{**base, **campos}]}


def testar_parse_de_partes_estruturadas():
    partes = mat._mat_parse_parties([
        ("João da Silva", "123.456.789-00"),
        ("Empresa XYZ Ltda", ""),
        ("", "12.345.678/0001-90"),
        ("", ""),                       # linha em branco não vira parte
        ("Fulano", "123"),              # documento inválido não vira documento
    ])
    assert [p["doc"] for p in partes] == ["12345678900", "", "12345678000190", ""]
    assert partes[1]["chave"] == "empresa xyz ltda"
    assert partes[3]["nome"] == "Fulano" and partes[3]["doc"] == ""


def testar_alerta_por_documento_continua_valendo():
    partes = mat._mat_parse_parties([("João da Silva", "123.456.789-00")])
    alertas = mat._mat_detectar_alertas(
        _transmissao(para_doc="12345678900"), partes, [], None
    )
    assert alertas == {"amarelo"}


def testar_alerta_sai_pelo_nome_quando_nao_ha_documento():
    partes = mat._mat_parse_parties([("Agropecuária Teste Ltda", "")])
    # Acento, caixa e pontuação diferentes; e a matrícula traz um sufixo a mais.
    alertas = mat._mat_detectar_alertas(
        _transmissao(de_nome="AGROPECUARIA TESTE LTDA - EPP"), partes, [], None
    )
    assert alertas == {"amarelo"}


def testar_nome_do_grupo_tambem_dispara():
    grupo = mat._mat_parse_parties([("Holdings XYZ S/A", "")])
    alertas = mat._mat_detectar_alertas(
        _transmissao(para_nome="Holdings XYZ S A"), [], grupo, None
    )
    assert alertas == {"amarelo"}


def testar_nome_curto_nao_casa_por_conteudo():
    """"Ana" não pode casar com "Ana Paula Rodrigues" — só igualdade exata."""
    partes = mat._mat_parse_parties([("Ana", "")])
    assert not mat._mat_detectar_alertas(
        _transmissao(de_nome="Ana Paula Rodrigues"), partes, [], None
    )
    assert mat._mat_detectar_alertas(_transmissao(de_nome="ANA"), partes, [], None)


def testar_nome_parecido_de_terceiro_nao_dispara():
    partes = mat._mat_parse_parties([("Agropecuária Teste Ltda", "")])
    assert not mat._mat_detectar_alertas(
        _transmissao(de_nome="Agropecuária Bandeirantes Ltda"), partes, [], None
    )


def testar_campos_achatados_viram_dois_blocos():
    """A UI manda nomes de devedor, docs de devedor, nomes do grupo, docs do grupo."""
    campos = ("Devedor 1", "Devedor 2", "111", "222",
              "Grupo 1", "Grupo 2", "333", "444")
    devedores, grupo = mat._mat_pares_dos_campos(campos)
    assert devedores == [("Devedor 1", "111"), ("Devedor 2", "222")]
    assert grupo == [("Grupo 1", "333"), ("Grupo 2", "444")]
    assert mat._mat_pares_dos_campos(()) == ([], [])
