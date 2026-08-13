"""Base de regras da prescrição intercorrente e a cronologia que a aplica.

O ponto do desenho: as REGRAS ficam escritas e versionadas, e o enquadramento por data
é feito em Python. O modelo só extrai fatos datados. Conta de data é exatamente o que
ele erra em silêncio.
"""
from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path

import gradio as gr
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


def testar_html_usa_id_proprio():
    """Duas cronologias na mesma página não podem disputar o mesmo elemento — a
    societária usa #timeline-export-area."""
    html_gerado = cronologia.render_html(cronologia.analisar(DADOS))
    assert "cronologia-prescricao-area" in html_gerado
    assert "timeline-export-area" not in html_gerado
    assert "Mov. 41" in html_gerado, "a referência processual acompanha o marco"


def testar_painel_e_so_a_linha_do_tempo():
    """O painel voltou a ser cronologia. A contagem continua existindo — em `avaliar` —,
    mas não é desenhada: o que se vê são os atos da execução."""
    html_gerado = cronologia.render_html(cronologia.analisar(DADOS))
    for sumido in ("presc-veredito", "Termo inicial", "Lapso de inércia",
                   "Regime aplicável", "Modelo de contagem"):
        assert sumido not in html_gerado, sumido
    assert cronologia.analisar(DADOS)["veredito"], "a conta continua sendo calculada"


def testar_sem_marcos_o_painel_orienta_em_vez_de_quebrar():
    assert "timeline-empty" in cronologia.render_html(cronologia.analisar({}))


def testar_edicao_no_painel_reaplica_o_enquadramento():
    """O painel é o editor: o JS devolve os marcos em "eventos" e o servidor refaz a
    conta com o título novo."""
    arvore = json.dumps({"titulo": "Nota promissória", "eventos": [
        {"data": "02/02/2023", "tipo": "arquivamento", "descricao": "Arquivo",
         "referencia": "Mov. 70"},
        {"data": "10/03/2014", "tipo": "distribuicao", "descricao": "Distribuída",
         "referencia": "fls. 1"},
        {"data": "", "tipo": "", "descricao": "", "referencia": ""},
    ], "_extra": {"vencimento_titulo": "01/01/2013"}}, ensure_ascii=False)
    _m, _p, dados, html_gerado, _b = cronologia.cronologia_aplicar_html("0", arvore)

    assert len(dados["marcos"]) == 2, "linha em branco não vira marco"
    assert dados["vencimento_titulo"] == "01/01/2013", "campo fora do painel sobrevive"
    assert "3 anos" in cronologia.analisar(dados)["prazo"]["prazo"],         "o prazo do título novo foi reaplicado no cálculo"
    assert html_gerado.index("10/03/2014") < html_gerado.index("02/02/2023"),         "o desenho sai em ordem cronológica"


def testar_painel_da_cronologia_e_editavel():
    html_gerado = cronologia.render_html(cronologia.analisar(DADOS))
    for marca in ("data-tl-raiz", "data-tl-evento", chr(39)+'data-tl-campo="titulo"'+chr(39),
                  chr(39)+'data-tl-campo="descricao"'+chr(39), "data-tl-remover-ato",
                  "+ Adicionar marco"):
        assert marca.strip(chr(39)) in html_gerado, marca


def testar_json_torto_da_cronologia_nao_derruba():
    for entrada in ("", "{", "[]", None):
        assert cronologia.aplicar_edicao_html(entrada) is None
    saida = cronologia.cronologia_aplicar_html("0", "{")
    assert saida[0] == "0" and isinstance(saida[2], type(gr.skip()))


def testar_exportar_html_e_autocontido():
    caminho = cronologia.exportar_html(DADOS)
    conteudo = Path(caminho).read_text(encoding="utf-8")
    assert conteudo.startswith("<!doctype html>")
    assert ".tl2-shell" in conteudo, "o CSS vai embutido, como no export da societária"
    assert "cronologia-prescricao-area" in conteudo


# ── Cálculo ──────────────────────────────────────────────────────────────────
# Sair do rótulo: quando começou a correr, quanto correu e se o lapso superou o prazo.

HOJE = date(2026, 8, 11)

MARCOS_CPC1973 = [
    {"data": "10/03/2012", "tipo": "distribuicao", "descricao": "Distribuída"},
    {"data": "20/09/2013", "tipo": "tentativa_infrutifera", "descricao": "Sem bens"},
    {"data": "11/11/2013", "tipo": "suspensao_921", "descricao": "Suspensa"},
]


def testar_iac1_sem_prazo_fixado_conta_um_ano_da_suspensao():
    """IAC 1/STJ: não havendo prazo fixado pelo juiz, o prazo começa um ano depois."""
    termo = pi.termo_inicial_intercorrente(MARCOS_CPC1973)
    assert termo["data"] == date(2014, 11, 11)
    assert "IAC 1" in termo["regra"] and "não houve prazo" in termo["regra"]


def testar_iac1_respeita_o_prazo_fixado_pelo_juiz():
    marcos = [dict(m, duracao="6 meses") if m["tipo"] == "suspensao_921" else m
              for m in MARCOS_CPC1973]
    termo = pi.termo_inicial_intercorrente(marcos)
    assert termo["data"] == date(2014, 5, 11)
    assert "fixado pelo juiz" in termo["regra"]


def testar_lei_14195_conta_da_primeira_tentativa_infrutifera():
    marcos = [{"data": "10/06/2022", "tipo": "tentativa_infrutifera", "descricao": "Sisbajud"},
              {"data": "01/09/2022", "tipo": "suspensao_921", "descricao": "Suspensa"}]
    termo = pi.termo_inicial_intercorrente(marcos)
    assert termo["data"] == date(2022, 6, 10), "no regime novo o termo antecipa"
    assert "14.195" in termo["regra"]


def testar_cpc2015_original_conta_do_fim_do_ano_de_suspensao():
    marcos = [{"data": "05/05/2018", "tipo": "suspensao_921", "descricao": "Suspensa"}]
    termo = pi.termo_inicial_intercorrente(marcos)
    assert termo["data"] == date(2019, 5, 5)
    assert "921" in termo["regra"]


def testar_sem_marco_de_inercia_nao_ha_termo_inicial():
    assert pi.termo_inicial_intercorrente(
        [{"data": "10/03/2012", "tipo": "distribuicao", "descricao": "x"}]) is None


# ── Transição do art. 2.028 ──────────────────────────────────────────────────

def testar_mais_da_metade_corrida_mantem_os_20_anos_do_cc1916():
    """Vencimento em 1990: em 11/01/2003 já haviam corrido 12,6 dos 20 anos."""
    prazos = pi.prazo_com_transicao("01/06/1990", "instrumento_particular")
    assert len(prazos) == 1
    assert prazos[0]["anos"] == 20
    assert prazos[0]["conta_de"] == date(1990, 6, 1)
    assert "CC/1916" in prazos[0]["regra"] and "art. 2.028" in prazos[0]["regra"]


def testar_menos_da_metade_corrida_aplica_o_cc2002_de_2003():
    """Vencimento em 2000: só 2,6 anos corridos — vale o prazo novo, de 11/01/2003."""
    prazos = pi.prazo_com_transicao("01/06/2000", "instrumento_particular")
    assert prazos[0]["anos"] == 5
    assert prazos[0]["conta_de"] == pi.VIGENCIA_CC_2002
    assert "menos da metade" in prazos[0]["regra"]


def testar_vencimento_depois_de_2003_nao_passa_pela_transicao():
    prazos = pi.prazo_com_transicao("10/01/2010", "instrumento_particular")
    assert prazos[0]["anos"] == 5 and prazos[0]["conta_de"] == date(2010, 1, 10)
    assert "art. 2.028" not in prazos[0]["regra"]


def testar_ccb_gera_dois_cenarios():
    prazos = pi.prazo_com_transicao("10/01/2010", "ccb")
    assert sorted(p["anos"] for p in prazos) == [3, 5]


# ── Lapsos de inércia ────────────────────────────────────────────────────────

def testar_penhora_zera_o_acumulado():
    marcos = MARCOS_CPC1973 + [
        {"data": "01/03/2020", "tipo": "penhora", "descricao": "Penhora do imóvel"}]
    resultado = pi.avaliar(marcos, "CCB", vencimento="10/01/2010", hoje=HOJE)
    corrido = resultado["cenarios"][0]["corrido"]
    assert corrido < 7, "a contagem recomeça da penhora, não do termo inicial"
    assert any("zera" in t["motivo"] for t in resultado["cenarios"][0]["trechos"])


def testar_recuperacao_judicial_pausa_sem_zerar():
    marcos = MARCOS_CPC1973 + [
        {"data": "01/01/2016", "tipo": "recuperacao_judicial", "descricao": "RJ deferida"},
        {"data": "01/01/2020", "tipo": "retomada", "descricao": "Execução retomada"}]
    trechos = pi.lapsos_de_inercia(marcos, date(2014, 11, 11), HOJE)
    pausados = [t for t in trechos if not t["corre"] and t["de"] != t["ate"]]
    assert pausados, "o período de RJ não pode correr contra o credor"


def testar_veredito_consumado_quando_o_lapso_supera_o_prazo():
    resultado = pi.avaliar(MARCOS_CPC1973, "CCB", vencimento="10/01/2010", hoje=HOJE)
    for cenario in resultado["cenarios"]:
        assert cenario["consumado"], f"{cenario['anos']} anos deveriam estar superados"
        assert cenario["data_limite"] == pi._somar_anos(date(2014, 11, 11), cenario["anos"])


def testar_extincao_encerra_a_contagem():
    marcos = MARCOS_CPC1973 + [
        {"data": "01/06/2016", "tipo": "extincao", "descricao": "Execução extinta"}]
    resultado = pi.avaliar(marcos, "CCB", vencimento="10/01/2010", hoje=HOJE)
    assert resultado["fim"] == date(2016, 6, 1)
    assert resultado["cenarios"][0]["corrido"] < 2


def testar_meses_da_duracao():
    assert pi.meses_da_duracao("1 ano") == 12
    assert pi.meses_da_duracao("6 meses") == 6
    assert pi.meses_da_duracao("90 dias") == 3
    assert pi.meses_da_duracao("sem prazo") is None


# ── Marcas de vigência e edição no painel ────────────────────────────────────

def _sem_css(marcacao: str) -> str:
    return re.sub(r"<style.*?</style>", "", marcacao, flags=re.DOTALL)


def testar_vigencia_de_lei_e_marca_na_trilha_e_nao_ato_do_processo():
    """Entrada em vigor não é fato dos autos: fica ENTRE as colunas, nunca dentro de uma."""
    corpo = _sem_css(cronologia.render_html(cronologia.analisar(DADOS)))
    assert "CPC/2015 · em vigor 18/03/2016" in corpo
    assert "Lei 14.195/2021 · em vigor 27/08/2021" in corpo
    for coluna in re.findall(r"<article.*?</article>", corpo, flags=re.DOTALL):
        assert "presc-lei" not in coluna, "a marca vazou para dentro de um ato"


def testar_so_marca_a_lei_que_virou_no_meio_da_execucao():
    """Execução toda posterior a 2003: marcar o CC/2002 ali seria ruído."""
    corpo = cronologia.render_html(cronologia.analisar(DADOS))
    assert "CC/2002 ·" not in corpo


def testar_a_trilha_reserva_largura_propria_para_a_marca():
    """A grade deixou de ter coluna de tamanho único: ato tem 200px, marca tem 44px."""
    corpo = _sem_css(cronologia.render_html(cronologia.analisar(DADOS)))
    trilha = re.search(r'grid-template-columns:([^"]*)"', corpo).group(1).split()
    assert trilha == ["200px", "44px", "200px", "200px", "44px", "200px"]


def testar_numero_do_processo_e_editavel_e_volta_pelo_painel():
    corpo = cronologia.render_html(cronologia.analisar(DADOS))
    assert 'data-tl-campo="processo"' in corpo
    arvore = json.dumps({"titulo": "CCB", "processo": "9999999-88.2020.8.16.0001",
                         "eventos": [{"data": "10/03/2014", "tipo": "distribuicao",
                                      "descricao": "D", "referencia": "1"}]})
    assert cronologia.aplicar_edicao_html(arvore)["processo"] == "9999999-88.2020.8.16.0001"


def testar_processo_de_html_antigo_ainda_vem_do_extra():
    """Painel gerado antes de o nº do processo ser desenhado o mandava em data-tl-extra."""
    arvore = json.dumps({"titulo": "CCB", "eventos": [], "_extra": {"processo": "123"}})
    assert cronologia.aplicar_edicao_html(arvore)["processo"] == "123"


def testar_html_exportado_nao_da_para_editar():
    """No Space o painel é editor; o arquivo exportado é registro e não tem onde salvar."""
    corpo = _sem_css(Path(cronologia.exportar_html(DADOS)).read_text(encoding="utf-8"))
    assert "data-tl-" not in corpo
    assert "tl2-add" not in corpo and "tl2-rm" not in corpo
    assert "tl2-dica" not in corpo
    assert "Sisbajud negativo" in corpo, "o conteúdo continua inteiro"
