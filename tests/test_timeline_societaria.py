"""Timeline Societária: falha de credencial vira mensagem acionável (não traceback)
e a exportação de imagem grava o PNG capturado do próprio HTML.

Contexto: quando o projeto da GEMINI_API_KEY_1 estourava o teto de gasto mensal, a aba
inteira caía com um traceback cru de `google.genai.errors.ClientError: 429`.
"""
from __future__ import annotations

import base64
import copy
import io
import json
import re
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
    assert saidas[-1][2]["empresa"] == "Teste Ltda"


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


# ── Edição no próprio painel ─────────────────────────────────────────────────
# O painel renderizado É o editor: contenteditable nos campos marcados com
# data-tl-campo e, ao concluir, o navegador devolve a árvore serializada. Não há mais
# grade com sintaxe dentro da célula, nem parse posicional deslocando campos.

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


def _arvore_do_evento(evento: dict) -> dict:
    """Reproduz o que o JS do navegador serializa a partir do painel."""
    montado = {campo: evento.get(campo, "") for campo in timeline.CAMPOS_TEXTO_DO_EVENTO}
    montado["_extra"] = {"categorias": evento.get("categorias", [])}
    for chave, campos in timeline.LISTAS_DO_EVENTO.items():
        montado[chave] = [{c: item.get(c, "") for c in campos} for item in evento.get(chave, [])]
    return montado


def _round_trip(dados: dict) -> dict:
    arvore = {"empresa": dados.get("empresa", ""), "cnpj": dados.get("cnpj", ""),
              "eventos": [_arvore_do_evento(e) for e in dados.get("eventos", [])]}
    return timeline.aplicar_edicao_html(json.dumps(arvore, ensure_ascii=False))


def testar_round_trip_nao_perde_nenhum_campo():
    resultado = _round_trip(DADOS_EDICAO)["eventos"][0]
    for chave, valor in EVENTO_COMPLETO.items():
        assert resultado[chave] == valor, f"campo {chave} mudou no round-trip"


def testar_campo_fora_do_painel_viaja_em_data_tl_extra():
    """"categorias" não é desenhado; viaja no atributo para o servidor não precisar do
    estado anterior — o que obrigaria a passar um gr.State pelo JavaScript."""
    assert _round_trip(DADOS_EDICAO)["eventos"][0]["categorias"] == ["Sócios", "Imóvel"]
    assert 'data-tl-extra="' in timeline.render_timeline_html(DADOS_EDICAO)


def testar_empresa_e_cnpj_sao_editaveis():
    arvore = {"empresa": "Outra Empresa Ltda", "cnpj": "99.999.999/0001-99",
              "eventos": [_arvore_do_evento(EVENTO_COMPLETO)]}
    novo = timeline.aplicar_edicao_html(json.dumps(arvore, ensure_ascii=False))
    assert novo["empresa"] == "Outra Empresa Ltda"
    assert novo["cnpj"] == "99.999.999/0001-99"


def testar_placeholder_nao_vira_conteudo():
    """O render escreve "—" e "Não informado." quando o campo está vazio. Voltando
    igual, tem que virar vazio de novo — senão o primeiro "Concluir edição" gravaria o
    travessão como se fosse o dado."""
    vazio = dict(EVENTO_COMPLETO, sede_apos="—", detalhamento="Não informado.",
                 numero_arquivamento="—", fonte="não informada")
    resultado = _round_trip({"empresa": "—", "cnpj": "—", "eventos": [vazio]})
    assert resultado["eventos"][0]["sede_apos"] == ""
    assert resultado["eventos"][0]["detalhamento"] == ""
    assert resultado["empresa"] == "" and resultado["cnpj"] == ""


def testar_linha_em_branco_nao_vira_item():
    arvore = {"empresa": "X", "cnpj": "", "eventos": [dict(
        _arvore_do_evento(EVENTO_COMPLETO),
        socios_apos=[{"nome": "Maria Souza", "participacao": "100%", "quotas": ""},
                     {"nome": "", "participacao": "", "quotas": ""}],
    )]}
    socios = timeline.aplicar_edicao_html(json.dumps(arvore))["eventos"][0]["socios_apos"]
    assert len(socios) == 1


def testar_ato_acrescentado_no_painel():
    """O JS acrescenta e remove colunas no DOM; o servidor recebe a lista resolvida."""
    arvore = {"empresa": "X", "cnpj": "", "eventos": [
        _arvore_do_evento(EVENTO_COMPLETO),
        dict({c: "" for c in timeline.CAMPOS_TEXTO_DO_EVENTO}, data="01/02/2024",
             ato="7ª Alteração", _extra={}),
    ]}
    novo = timeline.aplicar_edicao_html(json.dumps(arvore, ensure_ascii=False))
    assert len(novo["eventos"]) == 2
    assert novo["eventos"][1]["ato"] == "7ª Alteração"
    assert novo["eventos"][1]["socios_apos"] == []


def testar_json_torto_do_navegador_nao_derruba():
    for entrada in ("", "isso não é json", "[]", None, "{"):
        assert timeline.aplicar_edicao_html(entrada) is None


def testar_botao_alterna_entre_editar_e_concluir():
    # O JS devolve o modo NOVO: ao entrar manda "1"; ao concluir manda "0".
    entrando = timeline.timeline_aplicar_html("1", "")
    assert entrando[0] == "1" and "Concluir" in entrando[4]["value"]

    saindo = timeline.timeline_aplicar_html(
        "0", json.dumps({"empresa": "Nova", "cnpj": "", "eventos": []}))
    assert saindo[0] == "0" and "Editar" in saindo[4]["value"]
    assert saindo[2]["empresa"] == "Nova"


def testar_serializacao_torta_preserva_o_que_havia():
    """gr.skip() nas saídas de estado: JSON quebrado não pode zerar a timeline."""
    saida = timeline.timeline_aplicar_html("0", "{")
    assert saida[0] == "0"
    assert isinstance(saida[2], type(gr.skip()))
    assert isinstance(saida[3], type(gr.skip()))


def testar_marcacao_de_edicao_esta_no_html():
    html_gerado = timeline.render_timeline_html(DADOS_EDICAO)
    for marca in ("data-tl-raiz", "data-tl-evento", 'data-tl-lista="socios_apos"',
                  "data-tl-item", "data-tl-adicionar", "data-tl-remover",
                  'data-tl-campo="cnpj"', 'data-tl-campo="capital_social_anterior"'):
        assert marca in html_gerado, marca


def testar_ordenar_atos_por_data():
    fora_de_ordem = {"empresa": "X", "cnpj": "", "eventos": [
        dict(EVENTO_COMPLETO, data="05/06/2015", ato="Segundo"),
        dict(EVENTO_COMPLETO, data="10/03/2010", ato="Primeiro"),
    ]}
    ordenado, _html = timeline.ordenar_eventos(fora_de_ordem)
    assert [e["ato"] for e in ordenado["eventos"]] == ["Primeiro", "Segundo"]


def testar_codigo_orfao_foi_removido():
    """Eram três implementações paralelas do mesmo layout, mais o parse posicional."""
    for morto in ("timeline_exportar_imagem", "timeline_exportar_imagem_vertical",
                  "timeline_ver_tabela", "MOV_CORES", "data_to_rows", "rows_to_data",
                  "campos_do_evento", "selecionar_evento"):
        assert not hasattr(timeline, morto), f"{morto} deveria ter sido removido"


# ── Editar é do painel; exportar é registro ─────────────────────────────────
# A edição acontece no Space, onde há para onde salvar. O arquivo exportado sai sem
# nenhuma afordância: botão de + e ×, ganchos data-tl-* e a dica do modo de edição.

_DADOS_EXPORT = {"empresa": "Alfa Participações Ltda", "cnpj": "11.111.111/0001-11",
                 "eventos": [
                     {"data": "01/02/2010", "ato": "Constituição",
                      "socios_apos": [{"nome": "João", "participacao": "50%"}]},
                     {"data": "03/04/2015", "ato": "7ª Alteração",
                      "detalhamento": "Cessão de quotas para Maria.",
                      "socios_apos": [{"nome": "Maria", "participacao": "50%"}]},
                 ]}


def _sem_css(marcacao: str) -> str:
    return re.sub(r"<style.*?</style>", "", marcacao, flags=re.DOTALL)


def testar_html_exportado_nao_tem_como_editar():
    corpo = _sem_css(Path(timeline.timeline_exportar_html(_DADOS_EXPORT)).read_text(encoding="utf-8"))
    assert "data-tl-" not in corpo
    assert "tl2-add" not in corpo and "tl2-rm" not in corpo
    assert "tl2-dica" not in corpo


def testar_html_exportado_mantem_o_conteudo_e_o_pop_up():
    corpo = _sem_css(Path(timeline.timeline_exportar_html(_DADOS_EXPORT)).read_text(encoding="utf-8"))
    for esperado in ("Alfa Participações Ltda", "7ª Alteração", "Ver detalhamento",
                     "Cessão de quotas para Maria.", "Maria"):
        assert esperado in corpo, esperado


def testar_remover_edicao_nao_come_o_conteudo_das_caixas():
    """O × mora dentro da caixa; tirar o botão não pode levar a caixa junto."""
    limpo = timeline.remover_edicao(timeline.render_timeline_html(_DADOS_EXPORT))
    assert "tl2-move" in limpo and "Constituição" in limpo


def testar_o_ato_pode_ser_excluido_de_dentro_do_pop_up():
    corpo = timeline.render_timeline_html(_DADOS_EXPORT)
    assert corpo.count("data-tl-remover-ato") == 4, "um × na coluna e um botão no modal, por ato"
    assert "Excluir ato" in corpo
