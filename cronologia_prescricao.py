# -*- coding: utf-8 -*-
"""Cronologia processual com foco em prescrição intercorrente.

Divisão de trabalho deliberada: o modelo extrai apenas FATOS DATADOS do processo
(marcos, com data e referência). O enquadramento — qual regime do art. 921 governa cada
intervalo, quais leis entraram em vigor no meio do caminho, quanto tempo correu — é
calculado em Python a partir de `prescricao_intercorrente`. Conta de data é justamente o
que o modelo erra em silêncio e ninguém audita depois.

Reaproveita a linguagem visual da timeline societária (TL2_CSS), com prefixo próprio nos
ids: duas cronologias na mesma página não podem disputar o mesmo elemento.
"""

from __future__ import annotations

import html
import json
import os
import re
import tempfile
import unicodedata
from datetime import date

import gradio as gr
from google.genai import types

from prescricao_intercorrente import (
    MARCOS_LEGAIS,
    avaliar,
    TIPOS_DE_MARCO,
    _para_data,
    bloco_prompt,
    marcos_no_intervalo,
    prazo_do_titulo,
    regime_aplicavel,
)
from timeline_societaria import TL2_CSS, remover_edicao
from utils import _retry


# Cor por natureza do marco, na mesma paleta da timeline societária.
_CORES_MARCO = {
    "distribuicao": "#1A56A0",
    "citacao": "#235472",
    # Atividade do exequente em roxo: não é resultado (verde) nem frustração (laranja) —
    # é o que se opõe à inércia, e precisa se distinguir das duas coisas na leitura.
    "manifestacao_exequente": "#6B4C9A",
    "pedido_penhora": "#6B4C9A",
    "pedido_andamento": "#6B4C9A",
    "tentativa_infrutifera": "#DC4405",
    "suspensao_921": "#DC4405",
    "arquivamento": "#DC4405",
    "penhora": "#2F6B3A",
    "bens_localizados": "#2F6B3A",
    "parcelamento": "#2F6B3A",
    "embargos": "#1C6E8C",
    "recuperacao_judicial": "#1C6E8C",
    "retomada": "#2F6B3A",
    "extincao": "#A6486A",
    "outro": "#77817A",
}

_ROTULOS_MARCO = {
    "distribuicao": "Distribuição",
    "citacao": "Citação",
    "manifestacao_exequente": "Manifestação do exequente",
    "pedido_penhora": "Pedido de penhora",
    "pedido_andamento": "Pedido de andamento",
    "tentativa_infrutifera": "Tentativa infrutífera",
    "suspensao_921": "Suspensão (art. 921, III)",
    "arquivamento": "Arquivamento",
    "penhora": "Penhora efetivada",
    "bens_localizados": "Bens localizados",
    "parcelamento": "Parcelamento / acordo",
    "embargos": "Embargos / impugnação",
    "recuperacao_judicial": "Recuperação judicial",
    "retomada": "Retomada da execução",
    "extincao": "Extinção",
    "outro": "Outro ato",
}

# Rótulo → chave, para o painel aceitar a categoria escrita em português no modo de
# edição. Normalizado (sem acento, minúsculo) porque quem digita não repete o acento.
def _sem_acento(texto: str) -> str:
    cru = unicodedata.normalize("NFKD", str(texto or ""))
    return "".join(c for c in cru if not unicodedata.combining(c)).strip().lower()


_TIPOS_POR_ROTULO = {_sem_acento(rotulo): chave for chave, rotulo in _ROTULOS_MARCO.items()}
_TIPOS_POR_ROTULO.update({_sem_acento(chave): chave for chave in _ROTULOS_MARCO})

# As categorias reconhecidas, na dica do modo de edição: sem essa lista à vista, "clique
# para reescrever a categoria" não diz o que se pode escrever.
_CATEGORIAS_NA_DICA = " · ".join(html.escape(r) for r in _ROTULOS_MARCO.values())

PROMPT_MARCOS = """\
Extraia do material abaixo APENAS os marcos processuais que importam para a contagem da
prescrição intercorrente. Não conclua nada sobre prescrição — isso é feito depois, fora
daqui. Sua tarefa é localizar fatos datados e onde eles estão nos autos.

Marque cada fato com um destes tipos:
""" + "\n".join(f"- {t}" for t in TIPOS_DE_MARCO) + """

REGRAS
- Só registre fato que conste do material, com a data que consta. Não estime datas.
- Cada marco leva a referência processual exata (Mov./fls./Evento/ID).
- "tentativa_infrutifera" é a ciência de que não se localizou o devedor ou bens
  penhoráveis (diligência negativa, certidão do oficial, Sisbajud/Renajud sem resultado).
  Registre TODAS, principalmente a PRIMEIRA — ela é o termo inicial no regime atual.
- "suspensao_921" é a decisão que suspende a execução por não haver bens.
- Registre TODAS as manifestações do exequente, sem exceção e sem resumir várias numa
  só — cada petição é um marco, com a sua data. É por elas que se demonstra (ou se
  afasta) a inércia do credor, e uma que falte é um ano de silêncio aparente que não
  existiu. Use:
  · "pedido_penhora" — requerimento de penhora, arresto ou de diligência patrimonial
    (Sisbajud, Renajud, Infojud, CNIB, penhora no rosto dos autos, quebra de sigilo).
    É o PEDIDO; a penhora que se efetiva é "penhora".
  · "pedido_andamento" — petição que só pede impulso: prosseguimento, desarquivamento,
    reiteração de pedido não apreciado, cumprimento de decisão pendente.
  · "manifestacao_exequente" — qualquer outra petição do exequente (indicação de bens,
    atualização de cálculo, ciência de certidão negativa, resposta a intimação,
    substituição de CDA, pedido de IDPJ, habilitação de crédito).
- Não confunda o polo: manifestação do EXECUTADO, do Ministério Público ou de terceiro
  não entra nesses três tipos — se for relevante, use o tipo próprio ou "outro".
- Em "titulo", descreva o lastro exatamente como consta (ex.: "Cédula de Crédito
  Bancário nº 123", "Contrato de abertura de crédito"), com a referência.
- Em "vencimento_titulo", a data de vencimento do título, se constar.

Responda SOMENTE com JSON:
{
  "processo": "",
  "titulo": "",
  "vencimento_titulo": "DD/MM/AAAA ou vazio",
  "marcos": [
    {"data": "DD/MM/AAAA", "tipo": "", "descricao": "", "referencia": ""}
  ]
}

MATERIAL:
"""


# ── Extração ─────────────────────────────────────────────────────────────────

def extrair_marcos(texto_completo: str, relatorio: str, client, model: str) -> dict:
    material = ((relatorio or "").strip()[:200_000] + "\n\n" + (texto_completo or "").strip()[:500_000])
    prompt = bloco_prompt() + "\n" + PROMPT_MARCOS + material

    def _chamar():
        return client.models.generate_content(
            model=model,
            contents=[types.Content(role="user", parts=[types.Part(text=prompt)])],
            config=types.GenerateContentConfig(response_mime_type="application/json"),
        ).text

    bruto = _retry(_chamar, tentativas=3, espera_base=10)
    bruto = re.sub(r"^```[a-z]*\n?", "", (bruto or "").strip(), flags=re.IGNORECASE)
    bruto = re.sub(r"\n?```$", "", bruto.strip())
    dados = json.loads(bruto)
    if not isinstance(dados, dict):
        raise ValueError("A resposta da IA não contém um objeto JSON.")
    dados.setdefault("marcos", [])
    dados["marcos"] = ordenar_marcos([m for m in dados["marcos"] if isinstance(m, dict)])
    return dados


def ordenar_marcos(marcos: list) -> list:
    def chave(marco):
        quando = _para_data(marco.get("data"))
        return (quando or date(9999, 12, 31),)
    return sorted(marcos, key=chave)


# ── Enquadramento (determinístico) ───────────────────────────────────────────

def analisar(dados: dict) -> dict:
    """Sobrepõe aos marcos o regime de cada data e as viradas de lei entre eles."""
    marcos = ordenar_marcos([m for m in (dados or {}).get("marcos", []) if isinstance(m, dict)])
    prazo = prazo_do_titulo((dados or {}).get("titulo"))

    itens = []
    anterior = None
    for marco in marcos:
        quando = _para_data(marco.get("data"))
        entre = marcos_no_intervalo(anterior, quando) if anterior else []
        regime = regime_aplicavel(quando)
        itens.append({
            "marco": marco,
            "data": quando,
            "regime": regime,
            "leis_antes": entre,
        })
        if quando:
            anterior = quando

    # Termo inicial: a primeira tentativa infrutífera é o marco do regime atual; a
    # suspensão do art. 921 é o do regime anterior. Os dois são exibidos quando existem,
    # porque a diferença entre eles é o que decide o caso.
    def _primeira(tipo):
        for item in itens:
            if item["marco"].get("tipo") == tipo and item["data"]:
                return item
        return None

    return {
        "veredito": avaliar(marcos, (dados or {}).get("titulo"),
                            (dados or {}).get("vencimento_titulo")),
        "processo": (dados or {}).get("processo", ""),
        "titulo": (dados or {}).get("titulo", ""),
        "vencimento_titulo": (dados or {}).get("vencimento_titulo", ""),
        "prazo": prazo,
        "itens": itens,
        "primeira_tentativa": _primeira("tentativa_infrutifera"),
        "suspensao": _primeira("suspensao_921"),
        "leis": MARCOS_LEGAIS,
    }


# ── Render ───────────────────────────────────────────────────────────────────

def rotulo_do_marco(marco: dict) -> str:
    """A categoria como ela aparece na tela.

    Categoria fora da lista (escrita à mão no painel) fica gravada em "rotulo" e vence o
    rótulo padrão do tipo — é o que permite nomear um ato que a lista não previu sem
    perder a cor que a usuária escolheu.
    """
    livre = re.sub(r"\s+", " ", str(marco.get("rotulo") or "")).strip()
    tipo = str(marco.get("tipo") or "outro")
    return livre or _ROTULOS_MARCO.get(tipo, tipo)


def _coluna(item: dict) -> str:
    """Um ato da execução: data, categoria e a caixa com descrição e referência.

    Tudo o que se vê aqui é o dado bruto do marco — nada derivado —, então tudo é
    editável no próprio painel, a CATEGORIA inclusive: ela é desenhada com o rótulo em
    português e reconvertida em chave na volta (`_tipo_do_texto`). O tipo que a coluna já
    tinha viaja em data-tl-extra para servir de cor quando o texto digitado não
    corresponder a nenhuma categoria da lista.

    O enquadramento por regime saiu do card: ele agora aparece na linha do tempo como
    marca de vigência, entre os atos.
    """
    marco = item["marco"]
    tipo = str(marco.get("tipo") or "outro")
    cor = _CORES_MARCO.get(tipo, _CORES_MARCO["outro"])
    rotulo = rotulo_do_marco(marco)
    referencia = str(marco.get("referencia") or "")
    extra = html.escape(json.dumps({"tipo": tipo}, ensure_ascii=False), quote=True)

    return f"""
        <article class="tl2-col" data-tl-evento data-tl-extra="{extra}">
          <div class="tl2-col-head">
            <strong data-tl-campo="data" data-tl-rotulo="Data">{html.escape(str(marco.get("data") or "—"))}</strong>
            <span class="presc-natureza" data-tl-campo="tipo"
                  data-tl-rotulo="Categoria">{html.escape(rotulo)}</span>
            <button type="button" class="tl2-rm tl2-rm-ato" data-tl-remover-ato
                    title="Remover este marco">×</button>
          </div>
          <div class="tl2-axis"><span class="tl2-pin" style="background:{cor};box-shadow:0 0 0 1px {cor}"></span></div>
          <div class="tl2-moves">
            <div class="tl2-move" style="border-left-color:{cor};background:{cor}12">
              <span class="tl2-move-label" style="color:{cor}" data-tl-campo="tipo"
                    data-tl-rotulo="Categoria">{html.escape(rotulo)}</span>
              <strong data-tl-campo="descricao" data-tl-rotulo="Descrição">{html.escape(str(marco.get("descricao") or ""))}</strong>
              <small data-tl-campo="referencia" data-tl-rotulo="Referência">{html.escape(referencia)}</small>
            </div>
          </div>
        </article>
    """


def _marca_de_vigencia(lei: dict) -> str:
    """A entrada em vigor não é ato do processo — é uma marca na linha do tempo.

    Fica entre as colunas, como um divisor vertical: o que está à esquerda correu sob
    a lei anterior, o que está à direita, sob a nova. Não é editável de propósito —
    é data de vigência de lei, não fato dos autos.
    """
    nome = lei.get("curto") or lei["titulo"]
    return (
        '<div class="presc-lei" title="' + html.escape(lei["titulo"]) + '">'
        '<span class="presc-lei-seta">▼</span>'
        f'<span class="presc-lei-rot">{html.escape(nome)} · em vigor '
        f'{lei["data"].strftime("%d/%m/%Y")}</span></div>'
    )


_CSS_CRONOLOGIA = """
/* Natureza do ato, sob a data, fora do modo de edicao. */
.presc-natureza {
    display: block; margin-top: 2px; font-size: 10px; font-weight: 800; letter-spacing: .06em;
    text-transform: uppercase; color: #737670;
}

/* Marca de vigencia de lei: divisor vertical entre as colunas, nao um ato do processo. */
.presc-lei { position: relative; display: flex; flex-direction: column; align-items: center; }
.presc-lei::before {
    content: ""; position: absolute; top: 0; bottom: 0; left: 50%;
    border-left: 2px dashed #C9A227;
}
.presc-lei-seta {
    position: relative; font-size: 11px; line-height: 1; color: #C9A227;
}
.presc-lei-rot {
    position: relative; margin-top: 6px; padding: 9px 3px; writing-mode: vertical-rl;
    transform: rotate(180deg); background: #FBF6E7; border: 1px solid #E8D9A0;
    border-radius: 3px; font-size: 9px; font-weight: 800; letter-spacing: .07em;
    text-transform: uppercase; color: #8A6D1F; white-space: nowrap;
}
"""


def _extra_raiz(analise: dict) -> str:
    """Campos que o painel não desenha, para o round-trip não os perder."""
    return html.escape(json.dumps(
        {"vencimento_titulo": analise.get("vencimento_titulo", "")},
        ensure_ascii=False), quote=True)


def render_html(analise: dict) -> str:
    """A linha do tempo da execução — e só ela.

    A contagem (termo inicial, lapsos, prazo do título) continua sendo calculada em
    `analisar`, mas não é desenhada aqui: o painel é a cronologia dos atos. As três
    viradas de lei que mudam o regime aparecem como marca ENTRE as colunas, nunca
    dentro de um ato — vigência de lei não é fato dos autos.
    """
    itens = analise.get("itens") or []
    if not itens:
        return ('<div class="timeline-empty">Nenhum marco processual identificado. '
                "Gere a análise do processo primeiro.</div>")

    # A trilha alterna colunas de ato (200px) e marcas de vigência (44px), então a
    # largura vai explícita: a grade não tem mais uma coluna de tamanho único.
    celulas, larguras = [], []
    for item in itens:
        for lei in item["leis_antes"]:
            celulas.append(_marca_de_vigencia(lei))
            larguras.append("44px")
        celulas.append(_coluna(item))
        larguras.append("200px")

    return f"""
    <section class="tl2-shell" id="cronologia-prescricao-area" data-tl-raiz data-tl-extra="{_extra_raiz(analise)}">
      <style id="presc-export-style">{TL2_CSS}{_CSS_CRONOLOGIA}</style>
      <div class="tl2-head">
        <span class="tl2-kicker">Cronologia processual · prescrição intercorrente</span>
        <h2 data-tl-campo="titulo" data-tl-rotulo="Título / lastro">{html.escape(analise.get("titulo") or "Execução analisada")}</h2>
        <p>Execução <span data-tl-campo="processo" data-tl-rotulo="nº do processo">{html.escape(str(analise.get("processo") or "—"))}</span></p>
        <p class="tl2-dica">Modo de edição: clique em qualquer texto para reescrever —
           a <b>categoria</b> do ato inclusive, na linha sob a data ou no topo da caixa
           (as duas se atualizam juntas). Escreva uma destas para o robô entender e
           colorir: {_CATEGORIAS_NA_DICA}. Qualquer outro texto vira nome livre da caixa e
           a cor não muda. <b>×</b> remove o marco, <b>+ Adicionar marco</b> acrescenta um
           no fim; ao concluir, os marcos voltam ordenados por data.
           As marcas de vigência de lei saem das datas e não se editam.</p>
      </div>
      <div class="tl2-scroll">
        <div class="tl2-track" style="grid-template-columns:{" ".join(larguras)}">{"".join(celulas)}</div>
      </div>
      <button type="button" class="tl2-add tl2-add-ato" data-tl-adicionar-ato>+ Adicionar marco</button>
    </section>
    """


# ── Ponte com a interface ────────────────────────────────────────────────────







def exportar_html(dados: dict) -> str:
    """HTML autocontido, como o da timeline societária.

    Sai sem nenhuma afordância de edição: o arquivo exportado é o registro, e editar
    ali não teria para onde salvar. Quem edita é o painel no Space.
    """
    corpo = remover_edicao(render_html(analisar(dados or {})))
    documento = (
        "<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'>"
        "<title>Cronologia — prescrição intercorrente</title>"
        f"<style>body{{margin:0;padding:24px;background:#f5f4f1;"
        f"font-family:Arial,Helvetica,sans-serif}}{TL2_CSS}{_CSS_CRONOLOGIA}</style></head>"
        f"<body>{corpo}</body></html>"
    )
    caminho = os.path.join(tempfile.gettempdir(), "cronologia_prescricao.html")
    with open(caminho, "w", encoding="utf-8") as arquivo:
        arquivo.write(documento)
    return caminho


def _tipo_do_texto(texto: str, anterior: str = "") -> tuple[str, str]:
    """Resolve a categoria escrita no painel e devolve (tipo, rótulo livre).

    Aceita o rótulo em português ("Pedido de penhora"), a chave crua ("pedido_penhora") e
    qualquer uma das duas com acentuação ou caixa diferentes. Texto que não corresponde a
    nenhuma categoria da lista não é descartado nem rebaixado a "outro" em silêncio: fica
    como rótulo do marco, e a cor continua sendo a do tipo que a coluna já tinha — mudar
    a cor de um ato sem querer é pior do que não reconhecer o nome que se deu a ele.
    """
    normalizado = _sem_acento(texto)
    chave = _TIPOS_POR_ROTULO.get(normalizado)
    if chave:
        return chave, ""
    anterior = str(anterior or "") if str(anterior or "") in _ROTULOS_MARCO else "outro"
    if not normalizado:
        return anterior, ""
    return anterior, re.sub(r"\s+", " ", str(texto)).strip()


def aplicar_edicao_html(bruto: str):
    """Reconstrói os marcos a partir do painel editado.

    O JS é o mesmo da timeline societária e devolve a unidade repetida em "eventos";
    aqui cada uma delas é um marco. O título do lastro vem na raiz e é o que define o
    prazo aplicado, então mudá-lo refaz a conta inteira.
    """
    if not isinstance(bruto, str) or not bruto.strip():
        return None
    try:
        editado = json.loads(bruto)
    except json.JSONDecodeError:
        return None
    if not isinstance(editado, dict):
        return None

    marcos = []
    for recebido in editado.get("eventos") or []:
        if not isinstance(recebido, dict):
            continue
        marco = {campo: re.sub(r"\s+", " ", str(recebido.get(campo, "") or "")).strip()
                 for campo in ("data", "descricao", "referencia")}
        if marco["data"] in ("—", "-"):
            marco["data"] = ""
        # A categoria é desenhada pelo rótulo, então volta como rótulo: o tipo antigo,
        # que veio em data-tl-extra, é o fallback de cor.
        extra_col = recebido.get("_extra") if isinstance(recebido.get("_extra"), dict) else {}
        marco["tipo"], marco["rotulo"] = _tipo_do_texto(
            recebido.get("tipo"), extra_col.get("tipo"))
        # "tipo" fica fora do teste de vazio porque ele NUNCA vem vazio (sem categoria
        # reconhecida, cai em "outro"); quem diz se houve categoria é o texto recebido.
        if (marco["data"] or marco["descricao"] or marco["referencia"]
                or marco["rotulo"] or _sem_acento(recebido.get("tipo"))):
            marcos.append(marco)
    extra = editado.get("_extra") if isinstance(editado.get("_extra"), dict) else {}
    # O nº do processo é desenhado no cabeçalho desde que a contagem saiu do painel; o
    # data-tl-extra continua atendendo quem vier de um HTML gerado antes disso.
    processo = str(editado.get("processo", "") or extra.get("processo", "") or "").strip()
    return {"titulo": str(editado.get("titulo", "") or "").strip(), "marcos": marcos,
            "vencimento_titulo": str(extra.get("vencimento_titulo", "") or ""),
            "processo": "" if processo in ("—", "-") else processo}


def cronologia_aplicar_html(modo: str, bruto: str):
    """Liga e desliga a edição do painel da cronologia."""
    if str(modo or "0") == "1":
        return ("1", "", gr.skip(), gr.skip(),
                gr.update(value="Concluir edição", variant="primary"))

    novo = aplicar_edicao_html(bruto)
    if novo is None:
        return ("0", "", gr.skip(), gr.skip(),
                gr.update(value="Editar cronologia", variant="secondary"))
    # O vencimento e o nº do processo não são desenhados; vieram em data-tl-extra.
    return ("0", "", novo, render_html(analisar(novo)),
            gr.update(value="Editar cronologia", variant="secondary"))
