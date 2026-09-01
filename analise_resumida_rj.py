# -*- coding: utf-8 -*-
"""Análise Resumida da RJ — o resumo de uma tela, em tópicos.

Não é um terceiro checklist: é o recorte que se cola em e-mail ou anotação para
alguém entender o caso sem abrir o relatório. Sai sempre na mesma ordem —
recuperandos, advogados, administrador judicial, status, classe e valor do credor
analisado, lastros, garantias e ações relacionadas — porque a utilidade dele está
justamente em ser comparável entre casos.

Duas saídas do mesmo conteúdo: o texto em tópicos (para copiar) e o Word (para
anexar). A fonte é a mesma dos checklists — relatório consolidado como principal,
extração bruta só complementando.
"""

from __future__ import annotations

import re

from docx.shared import Cm

from checklist_rj import (
    _as_dict,
    _as_list,
    _as_str,
    _base_doc,
    _extrair,
    _montar_fonte_rj,
    _rodape_conf,
    _salvar,
    _titulo,
)
from dossie_ppa import (
    REGRA_CITACAO_PADRAO,
    REGRA_COMPLETUDE_PADRAO,
    _normalizar_caixa_alta,
    _para,
    _spacer,
    _TXT,
)
from legal_prompts import normalizar_referencias_objeto

# Marcador de "campo que quem analisa ainda vai preencher" — é como o modelo do
# resumo circula hoje, com asterisco no lugar do valor ainda não apurado.
PENDENTE = "*"

_MARCADORES = ["•", "∘", "▪"]
_RECUO_CM = 0.55


_PROMPT_RESUMO = REGRA_CITACAO_PADRAO + REGRA_COMPLETUDE_PADRAO + """
Você vai preencher a ANÁLISE RESUMIDA de uma Recuperação Judicial no formato JSON abaixo, a
partir do relatório consolidado e da extração bruta que vêm depois deste prompt. É um resumo
de uma tela: cada item é uma linha curta, sem parágrafo corrido e sem repetir o que já está
em outra linha.

═══ REGRA — PRIORIDADE DE FONTE ═══
O RELATÓRIO CONSOLIDADO é a fonte principal (já revisado, deduplicado e com as referências
corretas). O TEXTO BRUTO só complementa o que o relatório não cobrir. Divergindo, vale o
relatório.

═══ REGRA — CAMPO NÃO ENCONTRADO ═══
Campo sem informação nas duas fontes fica com string vazia "" (ou lista vazia). NÃO escreva
"Não consta", NÃO invente e NÃO preencha por dedução: aqui o vazio vira um marcador para quem
analisa completar depois.

═══ REGRA — REFERÊNCIA ═══
Este documento é um resumo de circulação rápida: NÃO anexe referência processual (fls./Mov./
ID/Evento) a nenhum campo, exceto se a referência fizer parte do próprio dado (ex.: número de
processo). As demais regras de fidelidade continuam valendo — nada que não esteja nas fontes.

═══ REGRA — CREDOR ANALISADO ═══
"credor" é o credor cujo crédito está sendo analisado — o exequente das execuções relacionadas
ou o titular do crédito discutido na impugnação. "classes" descreve a posição DESSE credor no
quadro de credores, uma entrada por classe em que ele figure, com o valor arrolado e a
representatividade percentual dele naquela classe quando o material trouxer. Se ele tiver
crédito extraconcursal, use "Extraconcursal" no campo "classe".

═══ REGRA — STATUS ═══
"status" é o estágio processual em uma linha (ex.: "Datas da AGC a serem designadas", "Plano
aprovado em AGC de DD/MM/AAAA", "Stay period até DD/MM/AAAA"). Uma entrada por fato relevante,
no máximo três.

═══ REGRA — GARANTIAS ═══
Em "garantias", só garantia real sobre bem IMÓVEL (hipoteca, alienação fiduciária de imóvel,
matrícula dada em garantia), uma entrada por imóvel, descrita em uma linha. Em
"discussao_essencialidade" da mesma entrada, informe se aquele imóvel é discutido como bem
essencial na RJ e como está essa discussão; se o material não tratar disso, deixe "".

═══ REGRA — AÇÕES RELACIONADAS ═══
"impugnacoes" traz apenas Incidentes de Impugnação de Crédito formalmente instaurados, com
número próprio. "execucoes" traz as execuções do credor analisado, com o SAT (saldo atualizado)
de cada uma quando o material informar. "outras_acoes" recebe qualquer outra ação relacionada
que não caiba nas duas listas anteriores.

Responda SOMENTE com o JSON.

{
  "rj_numero": "0000000-00.0000.0.00.0000",
  "grupo": "identificação curta do caso, como circula internamente (ex: Grupo Vieira)",
  "recuperandos": [{"nome": "", "documento": "CPF ou CNPJ"}],
  "advogados_recuperandos": [""],
  "administrador_judicial": [{"nome": "", "site": ""}],
  "status": [""],
  "credor": "Nome do credor analisado",
  "classes": [{"classe": "Classe II | Classe III | Extraconcursal | ...", "valor": "R$ ...", "representatividade": "~10%"}],
  "lastros": [""],
  "garantias": [{"descricao": "", "discussao_essencialidade": ""}],
  "impugnacoes": [{"numero": "", "polo_ativo": "", "finalidade": "", "status": ""}],
  "execucoes": [{"numero": "", "sat": "R$ ..."}],
  "outras_acoes": [{"tipo": "", "numero": "", "polo_ativo": "", "finalidade": "", "status": ""}]
}

"""


# ── Montagem do roteiro (nível, texto) ────────────────────────────────────────

def _txt(valor) -> str:
    return _normalizar_caixa_alta(str(valor or "").strip()).strip()


def _juntar(*partes, separador=" - ") -> str:
    return separador.join(p for p in (str(p or "").strip() for p in partes) if p)


def _bloco(linhas: list, rotulo: str, itens: list) -> None:
    """Um rótulo sempre aparece, ainda que sem item: o vazio é o lembrete de que
    aquilo continua para apurar — é assim que o resumo circula hoje."""
    linhas.append((0, f"{rotulo}:"))
    for item in itens or [""]:
        linhas.append((1, item))


def _linha_recuperando(item: dict) -> str:
    return _juntar(_txt(item.get("nome")), _txt(item.get("documento")))


def _linha_aj(item: dict) -> str:
    return _juntar(_txt(item.get("nome")), str(item.get("site") or "").strip(), separador=" | ")


def _linha_classe(item: dict) -> str:
    representatividade = str(item.get("representatividade") or "").strip()
    base = _juntar(_txt(item.get("classe")), _txt(item.get("valor")) or PENDENTE)
    return f"{base} ({representatividade})" if representatividade else base


def _classes(dados: dict) -> list:
    itens = [_linha_classe(c) for c in _as_list(dados.get("classes")) if _as_dict(c)]
    # A linha do extraconcursal é fixa no modelo: quando não há crédito fora do
    # concurso, ela fica com o marcador em vez de sumir.
    if not any(item.casefold().startswith("extraconcursal") for item in itens):
        itens.append(f"Extraconcursal - {PENDENTE}")
    return itens


def _garantias(linhas: list, dados: dict) -> None:
    linhas.append((0, "GARANTIAS:"))
    itens = [_as_dict(g) for g in _as_list(dados.get("garantias"))]
    if not itens:
        linhas.append((1, ""))
        return
    for ordem, garantia in enumerate(itens, 1):
        linhas.append((1, _juntar(f"{ordem}.", _txt(garantia.get("descricao")), separador=" ")))
        essencialidade = _txt(garantia.get("discussao_essencialidade"))
        linhas.append((1, _juntar(f"{ordem}.1. DISCUSSÃO DE ESSENCIALIDADE?",
                                  essencialidade, separador=" ")))


def _acoes_relacionadas(linhas: list, dados: dict) -> None:
    linhas.append((0, "AÇÕES RELACIONADAS:"))
    houve = False

    for impugnacao in (_as_dict(i) for i in _as_list(dados.get("impugnacoes"))):
        numero = _txt(impugnacao.get("numero"))
        if not numero and not _txt(impugnacao.get("finalidade")):
            continue
        houve = True
        linhas.append((1, _juntar("Impugnação de Crédito n.", numero, separador=" ")))
        for rotulo, chave in (("Polo ativo", "polo_ativo"), ("Finalidade", "finalidade"),
                              ("Status", "status")):
            linhas.append((2, f"{rotulo}: {_txt(impugnacao.get(chave))}"))

    execucoes = [_as_dict(e) for e in _as_list(dados.get("execucoes"))]
    execucoes = [e for e in execucoes if _txt(e.get("numero"))]
    if execucoes:
        houve = True
        linhas.append((1, "Execuções:"))
        for execucao in execucoes:
            sat = _txt(execucao.get("sat")) or PENDENTE
            linhas.append((2, f"{_txt(execucao.get('numero'))} – SAT: {sat}"))

    for acao in (_as_dict(a) for a in _as_list(dados.get("outras_acoes"))):
        numero = _txt(acao.get("numero"))
        tipo = _txt(acao.get("tipo")) or "Ação relacionada"
        if not numero and not _txt(acao.get("finalidade")):
            continue
        houve = True
        linhas.append((1, _juntar(f"{tipo} n.", numero, separador=" ")))
        for rotulo, chave in (("Polo ativo", "polo_ativo"), ("Finalidade", "finalidade"),
                              ("Status", "status")):
            linhas.append((2, f"{rotulo}: {_txt(acao.get(chave))}"))

    if not houve:
        linhas.append((1, ""))


def montar_roteiro(dados: dict) -> tuple:
    """Devolve (cabeçalho, [(nível, texto)]) — a mesma estrutura serve texto e Word."""
    dados = _as_dict(dados)
    cabecalho = _juntar(_as_str(dados.get("rj_numero")).strip(), _txt(dados.get("grupo")))

    linhas: list = []
    _bloco(linhas, "RECUPERANDOS",
           [_linha_recuperando(_as_dict(r)) for r in _as_list(dados.get("recuperandos"))])
    _bloco(linhas, "ADV",
           [_txt(a) for a in _as_list(dados.get("advogados_recuperandos")) if _txt(a)])
    _bloco(linhas, "AJ",
           [_linha_aj(_as_dict(a)) for a in _as_list(dados.get("administrador_judicial"))])
    _bloco(linhas, "STATUS", [_txt(s) for s in _as_list(dados.get("status")) if _txt(s)])
    _bloco(linhas, _juntar("CLASSE E VALOR", _txt(dados.get("credor"))), _classes(dados))
    _bloco(linhas, "LASTROS", [_txt(l) for l in _as_list(dados.get("lastros")) if _txt(l)])
    _garantias(linhas, dados)
    _acoes_relacionadas(linhas, dados)
    return cabecalho, linhas


# ── Saídas ────────────────────────────────────────────────────────────────────

def montar_texto(dados: dict) -> str:
    cabecalho, linhas = montar_roteiro(dados)
    corpo = [f"{'   ' * nivel}* {texto}".rstrip() for nivel, texto in linhas]
    return "\n".join([cabecalho, ""] + corpo)


def _build_resumo(dados: dict) -> str:
    cabecalho, linhas = montar_roteiro(dados)
    doc = _base_doc()
    _titulo(doc, "Análise Resumida — Recuperação Judicial")
    _para(doc, cabecalho, bold=True, size=11, color=_TXT, before=0, after=8,
          normalizar=False)
    # `montar_roteiro` já normalizou cada valor; aqui normalizar de novo só rebaixaria
    # os rótulos de estrutura ("RECUPERANDOS:") para capitalização de nome próprio.
    for nivel, texto in linhas:
        paragrafo = _para(doc, f"{_MARCADORES[min(nivel, len(_MARCADORES) - 1)]}  {texto}",
                          bold=(nivel == 0), size=9.5, color=_TXT,
                          before=(4 if nivel == 0 else 0), after=1, normalizar=False)
        paragrafo.paragraph_format.left_indent = Cm(_RECUO_CM * (nivel + 1))
    _spacer(doc)
    _rodape_conf(doc)
    return _salvar(doc, "Analise_Resumida_RJ", dados)


def gerar_analise_resumida(relatorio: str, texto_bruto: str, client, model: str,
                           credor: str = "") -> tuple:
    """Devolve (texto em tópicos, caminho do Word)."""
    fonte = _montar_fonte_rj(relatorio, texto_bruto)
    prompt = _PROMPT_RESUMO
    if str(credor or "").strip():
        prompt += (
            "CREDOR ANALISADO (informado por quem pediu a análise; use este e nenhum outro "
            f"no campo \"credor\" e nas classes): {str(credor).strip()}\n\n"
        )
    dados = _extrair(prompt, fonte, client, model)
    nomes_pdf = set(re.findall(r"(?:arquivo:\s*|^---\s*)([^\n—]+)", fonte, re.MULTILINE))
    dados = normalizar_referencias_objeto(dados, multiplos_pdfs=len(nomes_pdf) > 1)
    return montar_texto(dados), _build_resumo(dados)
