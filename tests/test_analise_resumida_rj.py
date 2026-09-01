"""Análise Resumida da RJ: o resumo de uma tela, sempre na mesma ordem.

A utilidade do documento está em ser comparável entre casos — quem lê procura o
mesmo item no mesmo lugar. Por isso os testes cobram a ORDEM e a presença dos
rótulos mesmo sem conteúdo, e não só "o dado apareceu".
"""
from __future__ import annotations

import json
import sys
import types as pytypes
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document  # noqa: E402

import analise_resumida_rj as resumida  # noqa: E402


DADOS = {
    "rj_numero": "0800472-89.2025.8.10.0081",
    "grupo": "GRUPO VIEIRA",
    "recuperandos": [
        {"nome": "EDUARDO VIEIRA", "documento": "632.923.191-53"},
        {"nome": "AGROPECUÁRIA ESTRELA DO XINGU LTDA", "documento": "03.907.502/0001-99"},
    ],
    "advogados_recuperandos": ["Aluizio Ramos Advogados Associados"],
    "administrador_judicial": [
        {"nome": "José Eduardo Pereira Junior", "site": "https://ejadvconsujus.com.br/"}
    ],
    "status": ["Datas da AGC a serem designadas"],
    "credor": "Sicredi Araxingu",
    "classes": [
        {"classe": "Classe II", "valor": "R$ 15.953.551,99", "representatividade": "~10%"},
        {"classe": "Classe III", "valor": "R$ 8.862.452,40", "representatividade": "~5%"},
    ],
    "lastros": ["CPR nº 123/2023"],
    "garantias": [
        {"descricao": "Hipoteca da matrícula 12.345 — Fazenda Boa Vista",
         "discussao_essencialidade": "Discutida como bem essencial, sem decisão"},
        {"descricao": "Alienação fiduciária da matrícula 6.789"},
    ],
    "impugnacoes": [{
        "numero": "0802810-36.2025.8.10.0081",
        "polo_ativo": "Cooperativa de Crédito, Poupança e Investimento do Araguaia e Xingu",
        "finalidade": "Requer a exclusão do crédito da RJ, por serem atos cooperativos",
        "status": "Pendente de julgamento",
    }],
    "execucoes": [
        {"numero": "1001231-74.2026.8.11.0059", "sat": "R$ 100.871,66"},
        {"numero": "1001458-64.2026.8.11.0059", "sat": "R$ 206.715,87"},
    ],
    "outras_acoes": [],
}

ROTULOS_NA_ORDEM = [
    "RECUPERANDOS:",
    "ADV:",
    "AJ:",
    "STATUS:",
    "CLASSE E VALOR - Sicredi Araxingu:",
    "LASTROS:",
    "GARANTIAS:",
    "AÇÕES RELACIONADAS:",
]


@pytest.fixture(scope="module")
def texto():
    return resumida.montar_texto(DADOS)


def _nivel_zero(texto: str) -> list:
    return [linha[2:] for linha in texto.split("\n") if linha.startswith("* ")]


def testar_cabecalho_traz_numero_e_grupo(texto):
    assert texto.split("\n")[0] == "0800472-89.2025.8.10.0081 - Grupo Vieira"
    assert texto.split("\n")[1] == "", "uma linha em branco separa o cabeçalho dos tópicos"


def testar_rotulos_saem_sempre_na_mesma_ordem(texto):
    assert _nivel_zero(texto) == ROTULOS_NA_ORDEM


def testar_recuo_marca_o_nivel_de_cada_item(texto):
    linhas = texto.split("\n")
    assert "   * Eduardo Vieira - 632.923.191-53" in linhas
    assert "   * Execuções:" in linhas
    assert "      * 1001231-74.2026.8.11.0059 – SAT: R$ 100.871,66" in linhas


def testar_nome_em_caixa_alta_vira_capitalizacao_natural(texto):
    """Mesma regra do resto da plataforma: nome de parte não sai todo em caixa alta."""
    assert "Agropecuária Estrela do Xingu Ltda - 03.907.502/0001-99" in texto
    assert "EDUARDO VIEIRA" not in texto


def testar_classe_traz_valor_e_representatividade(texto):
    assert "   * Classe II - R$ 15.953.551,99 (~10%)" in texto
    assert "   * Classe III - R$ 8.862.452,40 (~5%)" in texto


def testar_extraconcursal_aparece_mesmo_sem_credito_fora_do_concurso(texto):
    """A linha é fixa no modelo — sem crédito extraconcursal ela fica com o marcador."""
    assert f"   * Extraconcursal - {resumida.PENDENTE}" in texto


def testar_extraconcursal_informado_nao_vira_marcador():
    dados = dict(DADOS, classes=[{"classe": "Extraconcursal", "valor": "R$ 500.000,00"}])
    texto = resumida.montar_texto(dados)
    assert "   * Extraconcursal - R$ 500.000,00" in texto
    assert f"Extraconcursal - {resumida.PENDENTE}" not in texto


def testar_garantia_numerada_puxa_a_pergunta_de_essencialidade(texto):
    linhas = texto.split("\n")
    assert "   * 1. Hipoteca da matrícula 12.345 — Fazenda Boa Vista" in linhas
    assert ("   * 1.1. DISCUSSÃO DE ESSENCIALIDADE? Discutida como bem essencial, sem decisão"
            in linhas)
    # Sem resposta no material, a pergunta continua aberta para quem analisa.
    assert "   * 2.1. DISCUSSÃO DE ESSENCIALIDADE?" in linhas


def testar_acoes_relacionadas_separam_impugnacao_de_execucao(texto):
    linhas = texto.split("\n")
    assert "   * Impugnação de Crédito n. 0802810-36.2025.8.10.0081" in linhas
    assert "      * Status: Pendente de julgamento" in linhas
    assert "      * 1001458-64.2026.8.11.0059 – SAT: R$ 206.715,87" in linhas


def testar_execucao_sem_sat_fica_com_marcador():
    dados = dict(DADOS, execucoes=[{"numero": "1001231-74.2026.8.11.0059"}])
    assert f"– SAT: {resumida.PENDENTE}" in resumida.montar_texto(dados)


def testar_extracao_vazia_ainda_devolve_o_esqueleto():
    """Resumo sem dado nenhum continua sendo um roteiro do que falta apurar."""
    texto = resumida.montar_texto({})
    esperados = ["CLASSE E VALOR:" if r.startswith("CLASSE E VALOR") else r
                 for r in ROTULOS_NA_ORDEM]
    assert _nivel_zero(texto) == esperados
    assert texto.split("\n")[2:4] == ["* RECUPERANDOS:", "   *"]


def testar_tipos_malformados_nao_quebram_a_geracao():
    """Extração de PDF ruim às vezes devolve string onde se esperava lista."""
    bagunca = {"rj_numero": "1", "recuperandos": "texto solto", "classes": {"classe": "II"},
               "garantias": None, "execucoes": [None, "x"], "status": 7}
    texto = resumida.montar_texto(bagunca)
    assert "RECUPERANDOS:" in texto and "AÇÕES RELACIONADAS:" in texto


def testar_word_repete_o_mesmo_conteudo_do_texto():
    caminho = resumida._build_resumo(DADOS)
    assert Path(caminho).exists()
    conteudo = "\n".join(p.text for p in Document(caminho).paragraphs)
    assert "0800472-89.2025.8.10.0081 - Grupo Vieira" in conteudo
    for rotulo in ROTULOS_NA_ORDEM:
        assert rotulo in conteudo
    assert "1001231-74.2026.8.11.0059 – SAT: R$ 100.871,66" in conteudo


def testar_gerar_analise_resumida_liga_extracao_e_saidas():
    """Caminho completo com a IA dublada: devolve o texto e o caminho do Word."""
    prompts = []

    def _gerar(**kwargs):
        prompts.append(kwargs["contents"][0].parts[0].text)
        return pytypes.SimpleNamespace(text=json.dumps(DADOS, ensure_ascii=False))

    cliente = pytypes.SimpleNamespace(models=pytypes.SimpleNamespace(generate_content=_gerar))

    texto, caminho = resumida.gerar_analise_resumida(
        "relatório consolidado", "texto bruto", cliente, "modelo-x", credor="Sicredi Araxingu"
    )

    assert prompts, "a extração precisa passar pelo cliente Gemini"
    assert "CREDOR ANALISADO" in prompts[0], "credor informado na tela tem de chegar ao prompt"
    assert "relatório consolidado" in prompts[0]
    assert "RECUPERANDOS:" in texto
    assert Path(caminho).exists() and caminho.endswith(".docx")
