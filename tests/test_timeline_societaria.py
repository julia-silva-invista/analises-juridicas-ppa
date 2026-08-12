"""Timeline Societária: falha de credencial vira mensagem acionável (não traceback)
e a exportação de imagem grava o PNG capturado do próprio HTML.

Contexto: quando o projeto da GEMINI_API_KEY_1 estourava o teto de gasto mensal, a aba
inteira caía com um traceback cru de `google.genai.errors.ClientError: 429`.
"""
from __future__ import annotations

import base64
import copy
import io
import sys
import types as pytypes
from pathlib import Path

import gradio as gr
import pytest
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import timeline_societaria as timeline  # noqa: E402
import utils  # noqa: E402


ERRO_TETO_DE_GASTO = (
    "429 RESOURCE_EXHAUSTED. {'error': {'code': 429, 'message': 'Your project has exceeded "
    "its monthly spending cap. Please go to AI Studio at https://ai.studio/spend to manage "
    "your project spend cap.', 'status': 'RESOURCE_EXHAUSTED'}}"
)

JSON_RESPOSTA = '{"empresa": "Teste Ltda", "cnpj": "", "eventos": [{"data": "01/01/2020"}]}'


class _ArquivosFalsos:
    def __init__(self, registro, rotulo, erro):
        self._registro, self._rotulo, self._erro = registro, rotulo, erro

    def upload(self, file=None, **_):
        self._registro.append(self._rotulo)
        if self._erro:
            raise RuntimeError(self._erro)
        return pytypes.SimpleNamespace(
            name="files/x", uri="uri://x", mime_type="application/pdf",
            state=pytypes.SimpleNamespace(name="ACTIVE"),
        )

    def get(self, **_):
        return self.upload()

    def delete(self, **_):
        return None


class _ClienteFalso:
    """Dublê de genai.Client: falha no upload ou devolve um JSON pronto."""

    def __init__(self, registro, rotulo, erro=None, resposta=None, catalogo=()):
        self.files = _ArquivosFalsos(registro, rotulo, erro)
        self.models = pytypes.SimpleNamespace(
            generate_content=lambda **_: (
                registro.append(f"{rotulo}:gen"),
                pytypes.SimpleNamespace(text=resposta),
            )[1],
            list=lambda: [
                pytypes.SimpleNamespace(name=f"models/{n}", supported_actions=["generateContent"])
                for n in catalogo
            ],
        )


@pytest.fixture
def pdf_qualquer(tmp_path):
    caminho = tmp_path / "ato.pdf"
    caminho.write_bytes(b"%PDF-1.4\n%teste\n")
    return str(caminho)


@pytest.fixture(autouse=True)
def sem_espera(monkeypatch):
    """O _retry dorme entre tentativas — irrelevante para o teste."""
    monkeypatch.setattr(timeline.time, "sleep", lambda *_a, **_k: None)


def testar_teto_de_gasto_em_todas_as_chaves_vira_mensagem_acionavel(monkeypatch, pdf_qualquer):
    registro = []
    monkeypatch.setattr(timeline, "_get_gemini_clients", lambda: [
        _ClienteFalso(registro, "k1", erro=ERRO_TETO_DE_GASTO),
        _ClienteFalso(registro, "k2", erro=ERRO_TETO_DE_GASTO),
    ])

    saidas = []
    with pytest.raises(gr.Error) as capturado:
        for saida in timeline.timeline_analisar([pdf_qualquer]):
            saidas.append(saida)

    assert registro == ["k1", "k2"], "as duas credenciais têm que ser tentadas"
    assert "teto de gasto" in str(capturado.value)
    assert "ai.studio/spend" in str(capturado.value)
    assert "Traceback" not in str(capturado.value)
    # O painel também explica a falha, em vez de ficar preso no "carregando".
    assert "teto de gasto" in saidas[-1][0]
    assert "timeline-empty" in saidas[-1][1]


def testar_failover_conclui_na_segunda_chave(monkeypatch, pdf_qualquer):
    registro = []
    monkeypatch.setattr(timeline, "_get_gemini_clients", lambda: [
        _ClienteFalso(registro, "k1", erro=ERRO_TETO_DE_GASTO),
        _ClienteFalso(registro, "k2", resposta=JSON_RESPOSTA),
    ])

    saidas = list(timeline.timeline_analisar([pdf_qualquer]))

    assert registro == ["k1", "k2", "k2:gen"]
    assert "1 evento" in saidas[-1][0]
    assert saidas[-1][3]["empresa"] == "Teste Ltda"


def testar_404_nao_acusa_o_modelo_quando_ele_existe_na_chave(monkeypatch, pdf_qualquer):
    """O 404 pode vir da Files API, não do modelo. Mandar mexer em
    GEMINI_MODEL_TIMELINE nesse caso desperdiça o tempo de quem lê o aviso."""
    registro = []
    erro_404 = "404 NOT_FOUND. {'error': {'code': 404, 'message': 'File not found.'}}"
    monkeypatch.setattr(timeline, "_get_gemini_clients", lambda: [
        _ClienteFalso(registro, "k1", erro=erro_404,
                      catalogo=[timeline.MODEL_TIMELINE, "gemini-3.6-flash"]),
    ])

    with pytest.raises(gr.Error) as capturado:
        list(timeline.timeline_analisar([pdf_qualquer]))

    aviso = str(capturado.value)
    assert "GEMINI_MODEL_TIMELINE" not in aviso
    assert "o modelo existe nesta chave" in aviso
    assert "File not found" in aviso, "a resposta crua da API precisa aparecer"


def testar_404_lista_os_modelos_quando_o_configurado_nao_existe(monkeypatch, pdf_qualquer):
    registro = []
    erro_404 = "404 NOT_FOUND. {'error': {'code': 404, 'message': 'models/x is not found.'}}"
    monkeypatch.setattr(timeline, "_get_gemini_clients", lambda: [
        _ClienteFalso(registro, "k1", erro=erro_404, catalogo=["gemini-3.6-flash", "gemini-3.5-flash"]),
    ])

    with pytest.raises(gr.Error) as capturado:
        list(timeline.timeline_analisar([pdf_qualquer]))

    aviso = str(capturado.value)
    assert "GEMINI_MODEL_TIMELINE" in aviso
    assert "gemini-3.6-flash" in aviso and "gemini-3.5-flash" in aviso


def testar_mensagem_mostra_o_que_cada_chave_respondeu(monkeypatch, pdf_qualquer):
    """Só o erro da última credencial propaga; uma chave inválida no começo da fila
    precisa aparecer no aviso, senão o diagnóstico sai pela metade."""
    registro = []
    monkeypatch.setattr(timeline, "_get_gemini_clients", lambda: [
        _ClienteFalso(registro, "k1", erro="401 UNAUTHENTICATED. {'error': {'message': 'API key not valid'}}"),
        _ClienteFalso(registro, "k2", erro=ERRO_TETO_DE_GASTO),
    ])

    with pytest.raises(gr.Error) as capturado:
        list(timeline.timeline_analisar([pdf_qualquer]))

    aviso = str(capturado.value)
    assert "chave 1: 401" in aviso and "API key not valid" in aviso
    assert "chave 2: 429" in aviso
    assert "teto de gasto" in aviso


def testar_sem_arquivo_avisa_em_vez_de_quebrar():
    with pytest.raises(gr.Error, match="ao menos um ato"):
        list(timeline.timeline_analisar([]))


def testar_retry_nao_insiste_no_teto_de_gasto():
    """Teto de gasto não passa com espera: repetir só atrasa a mensagem de erro."""
    tentativas = []

    def _sempre_estourado():
        tentativas.append(1)
        raise RuntimeError(ERRO_TETO_DE_GASTO)

    with pytest.raises(RuntimeError):
        utils._retry(_sempre_estourado, tentativas=5, espera_base=1)

    assert len(tentativas) == 1


def testar_exportar_imagem_grava_a_captura_do_html():
    buffer = io.BytesIO()
    Image.new("RGB", (40, 30), "red").save(buffer, "PNG")
    captura = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode()

    caminho = timeline.timeline_salvar_imagem(captura)

    assert Image.open(caminho).size == (40, 30)


@pytest.mark.parametrize("captura", ["", None, "nao-e-data-url"])
def testar_exportar_imagem_sem_timeline_na_tela_avisa(captura):
    with pytest.raises(gr.Error, match="Não há timeline na tela"):
        timeline.timeline_salvar_imagem(captura)


def testar_exportar_imagem_com_base64_quebrado_avisa():
    with pytest.raises(gr.Error, match="não conseguiu gerar a imagem"):
        timeline.timeline_salvar_imagem("data:image/png;base64,!!!invalido!!!")


def testar_captura_js_recebe_um_argumento_e_devolve_lista():
    """O JS roda como pré-processamento do clique: precisa aceitar o valor atual do
    campo-ponte e devolver uma lista com os argumentos de timeline_salvar_imagem."""
    fonte = (Path(__file__).resolve().parents[1] / "app.py").read_text(encoding="utf-8")
    trecho = fonte.split('_CAPTURAR_TIMELINE_JS = """', 1)[1].split('"""', 1)[0]

    assert trecho.lstrip().startswith("async (_captura) =>")
    assert 'const vazio = [""];' in trecho
    assert 'return [canvas.toDataURL("image/png")];' in trecho
    assert 'return "";' not in trecho, "todo caminho de falha devolve a lista vazia"
    # Um evento só: encadear backend depois de um evento só-JS depende de comportamento
    # que varia entre a versão de desenvolvimento e a do Space.
    assert ".then(" not in fonte.split("tl_exportar_img_btn.click(", 1)[1].split(")\n", 1)[0]


# ── Edição por evento ────────────────────────────────────────────────────────
# A grade de 12 colunas achatava cinco listas do JSON numa mini-linguagem por célula
# ("Nome | 50%; Outro | 40%", "Cedente > Cessionário | % | valor"). Os joins omitiam
# campos vazios e os parses liam por posição, então um imóvel sem cartório voltava com
# o valor no lugar do cartório e o movimento trocado pelo default "integralizacao".

EVENTO_COMPLETO = {
    "data": "05/06/2015",
    "ato": "2ª Alteração",
    "numero_arquivamento": "20155551234",
    "detalhamento": "Cessão de quotas e integralização de imóvel.",
    "categorias": ["Sócios", "Imóvel"],
    "socios_apos": [{"nome": "Maria Souza", "participacao": "100%", "quotas": "100.000"}],
    "administradores_apos": [{"nome": "Maria Souza", "cargo": "Administradora"}],
    "capital_social_anterior": "R$ 100.000,00",
    "capital_social_apos": "R$ 1.000.000,00",
    "sede_apos": "Londrina/PR",
    "objeto_apos": "Agropecuária e holding",
    "cessoes": [{"cedente": "João da Silva", "cessionario": "Maria Souza",
                 "participacao": "50%", "valor": "R$ 1,00", "observacao": "valor simbólico"}],
    "imoveis": [{"matricula": "12.345", "cartorio": "1º CRI Curitiba", "cidade": "Curitiba",
                 "valor": "R$ 900.000,00", "movimento": "integralizacao",
                 "descricao": "Fazenda São Jorge"}],
    "filiais_apos": [{"nome": "Filial Maringá", "local": "Maringá/PR"}],
    "filiais_adicionadas": [{"nome": "Filial Maringá", "local": "Maringá/PR"}],
    "fonte": "2ª alteração, p. 1-3",
}

DADOS_EDICAO = {"empresa": "Agropecuária Teste Ltda", "cnpj": "12.345.678/0001-90",
                "eventos": [EVENTO_COMPLETO]}


def _round_trip(dados, indice=0):
    """Carrega o evento no editor e aplica de volta sem alterar nada."""
    campos = timeline.campos_do_evento(dados, indice)
    novo, _html, _rotulos = timeline.aplicar_edicao(dados, indice, *campos)
    return novo


def testar_round_trip_nao_perde_nenhum_campo():
    resultado = _round_trip(DADOS_EDICAO)["eventos"][0]
    for chave, valor in EVENTO_COMPLETO.items():
        assert resultado[chave] == valor, f"campo {chave} mudou no round-trip"


def testar_campo_fora_do_editor_sobrevive():
    """"categorias" era zerada a cada edição."""
    assert _round_trip(DADOS_EDICAO)["eventos"][0]["categorias"] == ["Sócios", "Imóvel"]


def testar_corrigir_a_data_nao_apaga_nada():
    """A reconciliação antiga era pela chave "data|ato": mudar a data fazia o lookup
    falhar e apagava capital_social_anterior e filiais_adicionadas — e, com elas, o
    card "Filial aberta" sumia da timeline sem explicação."""
    campos = timeline.campos_do_evento(DADOS_EDICAO, 0)
    campos[0] = "06/06/2015"          # data
    campos[1] = "2ª Alteração Contratual"  # ato
    evento = timeline.aplicar_edicao(DADOS_EDICAO, 0, *campos)[0]["eventos"][0]

    assert evento["data"] == "06/06/2015"
    assert evento["capital_social_anterior"] == "R$ 100.000,00"
    assert evento["filiais_adicionadas"] == [{"nome": "Filial Maringá", "local": "Maringá/PR"}]


def testar_imovel_sem_cartorio_nao_desloca_campos():
    """O caso que corrompia: campo vazio no meio da lista."""
    dados = copy.deepcopy(DADOS_EDICAO)
    dados["eventos"][0]["imoveis"] = [{"matricula": "12.345", "cartorio": "", "cidade": "",
                                       "valor": "R$ 800.000,00", "movimento": "saida",
                                       "descricao": ""}]
    imovel = _round_trip(dados)["eventos"][0]["imoveis"][0]
    assert imovel["valor"] == "R$ 800.000,00", "o valor escorregou para outra coluna"
    assert imovel["movimento"] == "saida", "o movimento virou o default silencioso"
    assert imovel["cartorio"] == ""


def testar_cessao_sem_participacao_preserva_o_valor():
    dados = copy.deepcopy(DADOS_EDICAO)
    dados["eventos"][0]["cessoes"] = [{"cedente": "A", "cessionario": "B", "participacao": "",
                                       "valor": "R$ 1,00", "observacao": ""}]
    cessao = _round_trip(dados)["eventos"][0]["cessoes"][0]
    assert cessao["participacao"] == "" and cessao["valor"] == "R$ 1,00"


def testar_linha_em_branco_da_tabelinha_e_descartada():
    campos = timeline.campos_do_evento(DADOS_EDICAO, 0)
    campos[9] = [["Maria Souza", "100%", "100.000"], ["", "", ""], [None, None, None]]
    socios = timeline.aplicar_edicao(DADOS_EDICAO, 0, *campos)[0]["eventos"][0]["socios_apos"]
    assert socios == [{"nome": "Maria Souza", "participacao": "100%", "quotas": "100.000"}]


def testar_empresa_e_cnpj_agora_sao_editaveis():
    """Não havia nenhuma via de correção: apareciam no cabeçalho do HTML e da imagem."""
    dados, html_gerado = timeline.aplicar_cabecalho(DADOS_EDICAO, "Outra Empresa Ltda", "99.999.999/0001-99")
    assert dados["empresa"] == "Outra Empresa Ltda"
    assert dados["cnpj"] == "99.999.999/0001-99"
    assert "Outra Empresa Ltda" in html_gerado
    assert dados["eventos"] == DADOS_EDICAO["eventos"], "editar o cabeçalho não toca nos atos"


def testar_adicionar_remover_e_ordenar_atos():
    dados, _html, _sel, indice, *campos = timeline.adicionar_evento(DADOS_EDICAO)
    assert len(dados["eventos"]) == 2 and indice == 1
    assert campos[1] == "Novo ato", "o ato novo já entra selecionado no editor"

    dados2, _h, _s, indice2, *_ = timeline.remover_evento(dados, 1)
    assert len(dados2["eventos"]) == 1 and indice2 == 0

    fora_de_ordem = {"empresa": "X", "cnpj": "", "eventos": [
        dict(EVENTO_COMPLETO, data="05/06/2015", ato="Segundo"),
        dict(EVENTO_COMPLETO, data="10/03/2010", ato="Primeiro"),
    ]}
    ordenado, _h, _s, indice3, *_ = timeline.ordenar_eventos(fora_de_ordem, 0)
    assert [e["ato"] for e in ordenado["eventos"]] == ["Primeiro", "Segundo"]
    assert indice3 == 1, "o ato que estava aberto continua aberto depois de reordenar"


def testar_rotulos_do_seletor_identificam_o_ato():
    assert timeline.rotulos_dos_eventos(DADOS_EDICAO) == ["1. 05/06/2015 — 2ª Alteração"]


def testar_indice_invalido_nao_quebra():
    for indice in (None, -5, 99, "abc"):
        assert timeline.campos_do_evento(DADOS_EDICAO, indice)
        timeline.aplicar_edicao(DADOS_EDICAO, indice, *timeline.campos_do_evento(DADOS_EDICAO, 0))


def testar_renderizadores_pil_orfaos_foram_removidos():
    """Eram três implementações paralelas do mesmo layout; duas nunca eram chamadas."""
    for morto in ("timeline_exportar_imagem", "timeline_exportar_imagem_vertical",
                  "timeline_ver_tabela", "MOV_CORES"):
        assert not hasattr(timeline, morto), f"{morto} deveria ter sido removido"
