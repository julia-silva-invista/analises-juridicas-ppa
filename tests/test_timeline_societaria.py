"""Timeline Societária: falha de credencial vira mensagem acionável (não traceback)
e a exportação de imagem grava o PNG capturado do próprio HTML.

Contexto: quando o projeto da GEMINI_API_KEY_1 estourava o teto de gasto mensal, a aba
inteira caía com um traceback cru de `google.genai.errors.ClientError: 429`.
"""
from __future__ import annotations

import base64
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
