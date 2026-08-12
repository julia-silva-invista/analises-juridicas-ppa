"""Base de regras da prescrição intercorrente e a cronologia que a aplica.

O ponto do desenho: as REGRAS ficam escritas e versionadas, e o enquadramento por data
é feito em Python. O modelo só extrai fatos datados. Conta de data é exatamente o que
ele erra em silêncio.
"""
from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import cronologia_prescricao as cronologia  # noqa: E402
import prescricao_intercorrente as pi  # noqa: E402


# ── Enquadramento por data ───────────────────────────────────────────────────

@pytest.mark.parametrize("quando, esperado", [
    ("10/03/2010", "cpc1973"),
    ("17/03/2016", "cpc1973"),          # véspera do CPC/2015
    ("18/03/2016", "cpc2015_original"),  # entrada em vigor
    ("26/08/2021", "cpc2015_original"),  # véspera da Lei 14.195
    ("27/08/2021", "lei14195"),          # entrada em vigor
    ("01/01/2026", "lei14195"),
])
def testar_regime_nas_datas_de_corte(quando, esperado):
    assert pi.regime_aplicavel(quando)["id"] == esperado


def testar_regime_aceita_date_e_texto_com_ruido():
    assert pi.regime_aplicavel(date(2018, 5, 1))["id"] == "cpc2015_original"
    assert pi.regime_aplicavel("decisão de 01/05/2018 (Mov. 42)")["id"] == "cpc2015_original"
    assert pi.regime_aplicavel("sem data") is None


def testar_marcos_de_vigencia_entre_dois_atos():
    """É o que a cronologia mostra entre um ato e o seguinte."""
    titulos = [m["titulo"] for m in pi.marcos_no_intervalo("01/01/2015", "01/01/2022")]
    assert any("CPC/2015" in t for t in titulos)
    assert any("14.195" in t for t in titulos)
    assert pi.marcos_no_intervalo("01/01/2017", "01/01/2018") == []
    # A virada exatamente na data de entrada em vigor conta para o intervalo.
    assert len(pi.marcos_no_intervalo("17/03/2016", "18/03/2016")) == 1


# ── Prazo por título ─────────────────────────────────────────────────────────

def testar_prazo_reconhece_o_lastro_pelo_texto_livre():
    assert pi.prazo_do_titulo("Cédula de Crédito Bancário (CCB) nº 301.906.503")["id"] == "ccb"
    assert pi.prazo_do_titulo("Nota promissória vinculada")["id"] == "nota_promissoria"
    assert pi.prazo_do_titulo("Contrato de abertura de crédito")["id"] == "instrumento_particular"
    assert pi.prazo_do_titulo("") is None


def testar_ccb_registra_a_divergencia_em_vez_de_escolher():
    """É o título mais comum da carteira e o de prazo mais disputado: 2 anos de
    diferença decidem o caso."""
    ccb = pi.PRAZOS_POR_TITULO["ccb"]
    assert ccb["status"] == "a_revisar"
    assert len(ccb["divergencia"]) == 2
    assert any("3 anos" in p for p in ccb["divergencia"])
    assert any("5 anos" in p for p in ccb["divergencia"])


# ── Disciplina da base ───────────────────────────────────────────────────────

def _todas_as_entradas():
    return (list(pi.REGIMES) + list(pi.PRAZOS_POR_TITULO.values())
            + list(pi.MITIGANTES) + list(pi.TRANSICAO) + list(pi.MARCOS_LEGAIS))


def testar_toda_entrada_tem_fundamento_e_status():
    for entrada in _todas_as_entradas():
        rotulo = entrada.get("rotulo") or entrada.get("titulo")
        assert entrada.get("status") in ("confirmado", "a_revisar"), rotulo
        if "detalhe" not in entrada:  # MARCOS_LEGAIS descrevem a virada, não fundamentam
            assert entrada.get("fundamento"), f"{rotulo} sem fundamento"


def testar_escopo_declarado_exclui_fiscal_e_trabalhista():
    """Regimes distintos; misturá-los produziria enquadramento errado em silêncio."""
    assert "FORA DO ESCOPO" in pi.__doc__
    corpo = Path(pi.__file__).read_text(encoding="utf-8")
    assert "art. 11-A" in corpo and "Súmula 314" in corpo, "a exclusão precisa ser explícita"


def testar_bloco_do_prompt_leva_regras_e_marca_o_que_conferir():
    bloco = pi.bloco_prompt()
    assert "Súmula 150/STF" in bloco
    assert "[CONFERIR]" in bloco, "itens divergentes têm que chegar sinalizados ao modelo"
    assert "Não recorra a memória própria" in bloco
    assert "Nunca afirme prescrição consumada como fato" in bloco
    assert "diga qual data falta" in bloco, "faltando data, o modelo tem que dizer qual"
    for mitigante in pi.MITIGANTES:
        assert mitigante["rotulo"] in bloco


# ── Cronologia ───────────────────────────────────────────────────────────────

DADOS = {
    "processo": "0001234-56.2014.8.16.0014",
    "titulo": "Cédula de Crédito Bancário (CCB) nº 301.906.503 (fls. 12)",
    "marcos": [
        {"data": "10/03/2014", "tipo": "distribuicao", "descricao": "Execução distribuída",
         "referencia": "fls. 1"},
        {"data": "20/09/2017", "tipo": "tentativa_infrutifera", "descricao": "Sisbajud negativo",
         "referencia": "Mov. 41"},
        {"data": "11/11/2018", "tipo": "suspensao_921", "descricao": "Execução suspensa",
         "referencia": "Mov. 55"},
        {"data": "02/02/2023", "tipo": "arquivamento", "descricao": "Autos ao arquivo",
         "referencia": "Mov. 70"},
    ],
}


def testar_cada_marco_recebe_o_regime_da_sua_data():
    itens = cronologia.analisar(DADOS)["itens"]
    assert [i["regime"]["id"] for i in itens] == [
        "cpc1973", "cpc2015_original", "cpc2015_original", "lei14195",
    ]


def testar_viradas_de_lei_aparecem_entre_os_atos():
    itens = cronologia.analisar(DADOS)["itens"]
    assert itens[0]["leis_antes"] == [], "nada antes do primeiro marco"
    assert any("CPC/2015" in l["titulo"] for l in itens[1]["leis_antes"])
    assert any("14.195" in l["titulo"] for l in itens[3]["leis_antes"])


def testar_termo_inicial_candidato_e_a_primeira_tentativa_infrutifera():
    analise = cronologia.analisar(DADOS)
    assert analise["primeira_tentativa"]["marco"]["data"] == "20/09/2017"
    assert analise["suspensao"]["marco"]["data"] == "11/11/2018"


def testar_marcos_fora_de_ordem_sao_ordenados():
    embaralhado = dict(DADOS, marcos=list(reversed(DADOS["marcos"])))
    datas = [i["marco"]["data"] for i in cronologia.analisar(embaralhado)["itens"]]
    assert datas == ["10/03/2014", "20/09/2017", "11/11/2018", "02/02/2023"]


def testar_html_usa_id_proprio_e_mostra_a_divergencia():
    """Duas cronologias na mesma página não podem disputar o mesmo elemento — a
    societária usa #timeline-export-area."""
    html_gerado = cronologia.render_html(cronologia.analisar(DADOS))
    assert "cronologia-prescricao-area" in html_gerado
    assert "timeline-export-area" not in html_gerado
    assert "DIVERGENTE" in html_gerado and "conferir" in html_gerado
    assert "Mov. 41" in html_gerado, "a referência processual acompanha o marco"


def testar_sem_marcos_o_painel_orienta_em_vez_de_quebrar():
    assert "timeline-empty" in cronologia.render_html(cronologia.analisar({}))


def testar_edicao_reaplica_o_enquadramento_sem_reordenar_o_que_se_digita():
    linhas = [["02/02/2023", "arquivamento", "Autos ao arquivo", "Mov. 70"],
              ["10/03/2014", "distribuicao", "Execução distribuída", "fls. 1"],
              ["", "", "", ""]]
    dados, html_gerado = cronologia.aplicar_marcos({}, "Nota promissória", linhas)

    assert [m["data"] for m in dados["marcos"]] == ["02/02/2023", "10/03/2014"], \
        "a ordem digitada é preservada; quem ordena é o desenho"
    assert "3 anos" in html_gerado, "o prazo do título novo foi reaplicado"
    assert html_gerado.index("10/03/2014") < html_gerado.index("02/02/2023"), \
        "o desenho sai em ordem cronológica"


def testar_exportar_html_e_autocontido():
    caminho = cronologia.exportar_html(DADOS)
    conteudo = Path(caminho).read_text(encoding="utf-8")
    assert conteudo.startswith("<!doctype html>")
    assert ".tl2-shell" in conteudo, "o CSS vai embutido, como no export da societária"
    assert "cronologia-prescricao-area" in conteudo
