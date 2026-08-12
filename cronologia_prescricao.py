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
from timeline_societaria import TL2_CSS
from utils import _retry

COLUNAS_MARCOS = ["Data", "Tipo", "Descrição", "Referência"]

# Cor por natureza do marco, na mesma paleta da timeline societária.
_CORES_MARCO = {
    "distribuicao": "#1A56A0",
    "citacao": "#235472",
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

def _card(rotulo: str, titulo: str, extra: str, cor: str) -> str:
    linha_extra = f"<small>{html.escape(extra)}</small>" if extra else ""
    return (
        f'<div class="tl2-move" style="border-left-color:{cor};background:{cor}12">'
        f'<span class="tl2-move-label" style="color:{cor}">{html.escape(rotulo)}</span>'
        f"<strong>{html.escape(titulo)}</strong>{linha_extra}</div>"
    )


def _coluna(item: dict) -> str:
    marco = item["marco"]
    tipo = str(marco.get("tipo") or "outro")
    cor = _CORES_MARCO.get(tipo, _CORES_MARCO["outro"])
    rotulo = _ROTULOS_MARCO.get(tipo, tipo)

    cards = []
    for lei in item["leis_antes"]:
        cards.append(_card("Entrada em vigor", lei["titulo"],
                           lei["data"].strftime("%d/%m/%Y"), "#8A6D1F"))
    cards.append(_card(rotulo, str(marco.get("descricao") or rotulo),
                       str(marco.get("referencia") or ""), cor))

    regime = item.get("regime")
    rodape = (f'<div class="tl2-quadro">Regime aplicável<strong>'
              f'{html.escape(regime["rotulo"])}</strong></div>') if regime else ""

    return (
        '<article class="tl2-col">'
        '<div class="tl2-col-head">'
        f'<strong>{html.escape(str(marco.get("data") or "—"))}</strong>'
        f"<span>{html.escape(rotulo)}</span></div>"
        '<div class="tl2-axis"><span class="tl2-pin"></span></div>'
        f'<div class="tl2-moves">{"".join(cards)}</div>{rodape}</article>'
    )


def _painel_prazo(analise: dict) -> str:
    prazo = analise.get("prazo")
    if not prazo:
        titulo = analise.get("titulo") or "não identificado"
        return (f'<p class="tl2-fonte">Título: {html.escape(str(titulo))} — prazo não '
                "reconhecido pela base de regras. Informe o lastro para que o prazo seja aplicado.</p>")
    aviso = " ⚠️ prazo divergente — conferir" if prazo.get("status") == "a_revisar" else ""
    posicoes = "".join(f"<li>{html.escape(p)}</li>" for p in prazo.get("divergencia", []))
    lista = f"<ul>{posicoes}</ul>" if posicoes else ""
    return (
        f'<p class="tl2-fonte"><strong>{html.escape(prazo["rotulo"])}</strong> — '
        f'{html.escape(prazo["prazo"])} ({html.escape(prazo["termo"])}), '
        f'{html.escape(prazo["fundamento"])}{aviso}</p>{lista}'
    )


def _anos(valor: float) -> str:
    anos = int(valor)
    meses = round((valor - anos) * 12)
    if meses == 12:
        anos, meses = anos + 1, 0
    partes = []
    if anos:
        partes.append(f"{anos} ano" + ("s" if anos != 1 else ""))
    if meses:
        partes.append(f"{meses} mês" if meses == 1 else f"{meses} meses")
    return " e ".join(partes) or "menos de um mês"


def _linha_lapso(trecho: dict) -> str:
    if trecho["de"] == trecho["ate"]:
        return (f'<li class="presc-zera">{trecho["de"].strftime("%d/%m/%Y")} — '
                f'{html.escape(trecho["motivo"])}</li>')
    corre = "corre" if trecho["corre"] else f'não corre ({html.escape(trecho["motivo"])})'
    return (f'<li>{trecho["de"].strftime("%d/%m/%Y")} → {trecho["ate"].strftime("%d/%m/%Y")}'
            f' · {_anos(trecho["anos"])} · <b>{corre}</b>'
            f' · acumulado {_anos(trecho["acumulado"])}</li>')


def _painel_veredito(veredito: dict) -> str:
    """A conta, aberta: termo inicial, lapsos de inércia e resultado por cenário.

    O risco é apresentado como risco. "Lapso superior ao prazo" quer dizer que a conta
    fecha — não que a prescrição esteja declarada. Os mitigantes e a leitura dos autos
    continuam sendo do analista.
    """
    termo = veredito.get("termo")
    if not termo or not termo.get("data"):
        return ('<div class="presc-veredito"><h4>Contagem</h4><p>Não foi possível fixar o '
                'termo inicial: não há nos autos decisão de suspensão do art. 921, III nem '
                'ciência de tentativa infrutífera de localização. É a data que falta para '
                'fechar a conta.</p></div>')

    regime = termo.get("regime") or {}
    cabecalho = (
        f'<h4>Contagem</h4>'
        f'<p><b>Termo inicial:</b> {termo["data"].strftime("%d/%m/%Y")} — '
        f'{html.escape(termo["regra"])}'
        + (f' <i>({html.escape(regime.get("rotulo", ""))})</i>' if regime else "")
        + f' Marco de origem: {html.escape(str(termo["marco"].get("descricao") or ""))}'
        + (f' ({html.escape(str(termo["marco"].get("referencia") or ""))})'
           if termo["marco"].get("referencia") else "") + "</p>"
    )
    if veredito.get("extincao"):
        cabecalho += (f'<p><b>Contagem encerrada em</b> {veredito["fim"].strftime("%d/%m/%Y")}'
                      f' pela extinção registrada nos autos.</p>')

    blocos = []
    for cenario in veredito.get("cenarios", []):
        classe = "presc-consumado" if cenario["consumado"] else "presc-corrente"
        titulo = f'Prazo de {_anos(cenario["anos"])}'
        if cenario["consumado"]:
            situacao = (f'<b>Lapso de inércia de {_anos(cenario["corrido"])} — SUPERIOR ao '
                        f'prazo.</b> Pela contagem, o prazo se completaria em '
                        f'{cenario["data_limite"].strftime("%d/%m/%Y")}.')
        else:
            situacao = (f'Lapso de inércia de {_anos(cenario["corrido"])}; faltam '
                        f'{_anos(cenario["faltam"])}'
                        + (f', completando em {cenario["data_limite"].strftime("%d/%m/%Y")}'
                           if cenario.get("data_limite") else "") + ".")
        lapsos = "".join(_linha_lapso(t) for t in cenario.get("trechos", []))
        blocos.append(
            f'<div class="presc-cenario {classe}"><h5>{titulo}</h5>'
            f'<p class="presc-regra">{html.escape(cenario["regra"])}</p>'
            f"<p>{situacao}</p>"
            + (f'<ul class="presc-lapsos">{lapsos}</ul>' if lapsos else "")
            + "</div>"
        )

    aviso = ('<p class="tl2-fonte">Duas contas porque o prazo do título é divergente: o '
             'resultado muda conforme a posição adotada.</p>') if veredito.get("divergente") else ""
    return (f'<div class="presc-veredito">{cabecalho}{"".join(blocos)}{aviso}'
            '<p class="tl2-fonte">Modelo de contagem: penhora efetivada, localização de bens, '
            'parcelamento e retomada zeram o acumulado; recuperação judicial e embargos '
            'pausam sem zerar. Risco é risco — confira os mitigantes e a leitura dos autos '
            'antes de concluir.</p></div>')


_CSS_VEREDITO = """
.presc-veredito { margin-top: 16px; padding: 14px 16px; background: #fbfbfa;
    border: 1px solid #e4e4e0; }
.presc-veredito h4 { margin: 0 0 6px; font-size: 12px; text-transform: uppercase;
    letter-spacing: .08em; color: #1A56A0; }
.presc-veredito p { margin: 4px 0; font-size: 12px; color: #4a4d48; line-height: 1.6; }
.presc-cenario { margin-top: 10px; padding: 10px 12px; border-left: 3px solid #77817a;
    background: #fff; }
.presc-cenario h5 { margin: 0 0 4px; font-size: 12px; font-weight: 800; color: #1f211f; }
.presc-corrente { border-left-color: #2F6B3A; }
.presc-consumado { border-left-color: #DC4405; background: rgba(220, 68, 5, .05); }
.presc-regra { font-style: italic; color: #737670 !important; }
.presc-lapsos { margin: 6px 0 0; padding-left: 18px; font-size: 11px; color: #4a4d48;
    line-height: 1.7; }
.presc-lapsos .presc-zera { list-style: none; margin-left: -18px; color: #2F6B3A;
    font-weight: 700; }
"""


def render_html(analise: dict) -> str:
    itens = analise.get("itens") or []
    if not itens:
        return ('<div class="timeline-empty">Nenhum marco processual identificado. '
                "Gere a análise do processo primeiro.</div>")

    termo = analise.get("primeira_tentativa") or analise.get("suspensao")
    resumo = []
    if termo:
        resumo.append("Termo inicial candidato: "
                      + str(termo["marco"].get("data") or "")
                      + " — " + _ROTULOS_MARCO.get(termo["marco"].get("tipo", ""), ""))
    if analise.get("processo"):
        resumo.append(str(analise["processo"]))

    return f"""
    <section class="tl2-shell" id="cronologia-prescricao-area">
      <style id="presc-export-style">{TL2_CSS}{_CSS_VEREDITO}</style>
      <div class="tl2-head">
        <span class="tl2-kicker">Cronologia processual · prescrição intercorrente</span>
        <h2>{html.escape(analise.get("titulo") or "Execução analisada")}</h2>
        <p>{html.escape(" · ".join(resumo)) if resumo else ""}</p>
        {_painel_prazo(analise)}
        {_painel_veredito(analise.get("veredito") or {})}
      </div>
      <div class="tl2-scroll">
        <div class="tl2-track" style="--tl2-cols:{len(itens)}">{"".join(_coluna(i) for i in itens)}</div>
      </div>
    </section>
    """


# ── Ponte com a interface ────────────────────────────────────────────────────

def marcos_para_linhas(dados: dict) -> list[list[str]]:
    linhas = [
        [str(m.get("data", "") or ""), str(m.get("tipo", "") or ""),
         str(m.get("descricao", "") or ""), str(m.get("referencia", "") or "")]
        for m in (dados or {}).get("marcos", []) if isinstance(m, dict)
    ]
    return linhas or [["", "", "", ""]]


def linhas_para_marcos(linhas) -> list[dict]:
    if hasattr(linhas, "values"):
        linhas = linhas.values.tolist()
    marcos = []
    for linha in linhas or []:
        valores = [str(v).strip() if v is not None and str(v) != "nan" else "" for v in linha]
        valores += [""] * (4 - len(valores))
        if not any(valores):
            continue
        marcos.append({"data": valores[0], "tipo": valores[1],
                       "descricao": valores[2], "referencia": valores[3]})
    return marcos


def aplicar_marcos(dados: dict, titulo: str, linhas):
    """Reaplica o enquadramento a cada edição — é o preview ao vivo da cronologia.

    Não devolve as linhas de volta ao dataframe de propósito: reescrevê-lo a cada tecla
    reordenaria as linhas embaixo do cursor. A ordenação por data acontece só na
    montagem do desenho (`analisar`), não no que está sendo editado.
    """
    dados = dict(dados or {})
    dados["titulo"] = str(titulo or "").strip()
    dados["marcos"] = linhas_para_marcos(linhas)
    return dados, render_html(analisar(dados))


def exportar_html(dados: dict) -> str:
    """HTML autocontido, como o da timeline societária."""
    corpo = render_html(analisar(dados or {}))
    documento = (
        "<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'>"
        "<title>Cronologia — prescrição intercorrente</title>"
        f"<style>body{{margin:0;padding:24px;background:#f5f4f1;"
        f"font-family:Arial,Helvetica,sans-serif}}{TL2_CSS}{_CSS_VEREDITO}</style></head>"
        f"<body>{corpo}</body></html>"
    )
    caminho = os.path.join(tempfile.gettempdir(), "cronologia_prescricao.html")
    with open(caminho, "w", encoding="utf-8") as arquivo:
        arquivo.write(documento)
    return caminho
