# -*- coding: utf-8 -*-
"""Base de regras da prescrição intercorrente na EXECUÇÃO COMUM (cível/bancária).

Por que existe: pedir ao modelo que "lembre" o regime aplicável a cada intervalo é caro
e instável — a resposta muda entre execuções e não dá para auditar. Aqui as regras ficam
escritas uma vez, versionadas e revisáveis; o modelo só extrai FATOS DATADOS do processo
e o enquadramento por data é feito em Python (`regime_aplicavel`), deterministicamente.

FORA DO ESCOPO, de propósito: execução fiscal (art. 40 da LEF, Súmula 314/STJ, Tema
566/STJ) e execução trabalhista (art. 11-A da CLT, IN 41/2018 do TST). São árvores de
regra próprias; misturá-las aqui produziria enquadramento errado em silêncio.

⚠️ REVISÃO JURÍDICA: este arquivo é conteúdo jurídico curado, não código neutro. Cada
entrada declara `status`:
  - "confirmado" — dispositivo de lei com texto e vigência estáveis;
  - "a_revisar"  — depende de jurisprudência, tem divergência real, ou a redação exata
                   precisa ser conferida antes de virar fundamento de parecer.
Nada aqui deve ser citado em parecer sem conferência na fonte primária. Onde há
divergência, a entrada registra as DUAS posições em vez de escolher uma.
"""

from __future__ import annotations

import re
from datetime import date

# ── Marcos de vigência ───────────────────────────────────────────────────────
# Datas usadas para cortar a linha do tempo do processo em janelas de regime.

VIGENCIA_CC_2002 = date(2003, 1, 11)
VIGENCIA_CPC_2015 = date(2016, 3, 18)
VIGENCIA_LEI_14195 = date(2021, 8, 27)

MARCOS_LEGAIS = [
    {
        "data": VIGENCIA_CC_2002,
        "titulo": "Código Civil de 2002 entra em vigor",
        "curto": "CC/2002",
        "detalhe": "Novos prazos prescricionais. A transição é regida pelo art. 2.028: "
                   "permanece o prazo da lei anterior quando o novo o reduziu E já havia "
                   "transcorrido mais da metade do prazo antigo em 11/01/2003.",
        "status": "confirmado",
    },
    {
        "data": VIGENCIA_CPC_2015,
        "titulo": "CPC/2015 entra em vigor — art. 921 disciplina a intercorrente",
        "curto": "CPC/2015",
        "detalhe": "Pela primeira vez a execução comum ganha disciplina expressa: suspensão "
                   "de 1 ano sem correr prescrição (§1º), arquivamento (§2º) e, decorrido o "
                   "prazo sem localização de bens, curso da prescrição intercorrente (§4º), "
                   "reconhecível de ofício depois de ouvidas as partes (§5º).",
        "status": "confirmado",
    },
    {
        "data": VIGENCIA_LEI_14195,
        "titulo": "Lei 14.195/2021 altera o art. 921",
        "curto": "Lei 14.195/2021",
        "detalhe": "Muda o TERMO INICIAL: o prazo passa a contar da ciência da primeira "
                   "tentativa infrutífera de localizar o devedor ou bens penhoráveis, e não "
                   "do fim do ano de suspensão. Redação exata e alcance intertemporal devem "
                   "ser conferidos antes de fundamentar.",
        "status": "a_revisar",
    },
]


# ── Regimes, por janela de vigência ──────────────────────────────────────────

REGIMES = [
    {
        "id": "cpc1973",
        "rotulo": "CPC/1973 (até 17/03/2016)",
        "inicio": None,
        "fim": date(2016, 3, 17),
        "fundamento": "CPC/1973, art. 791, III",
        "regra": "Não havia disciplina legal expressa da prescrição intercorrente na "
                 "execução comum. O art. 791, III permitia suspender a execução sem prazo "
                 "determinado quando não encontrados bens penhoráveis.",
        "termo_inicial": "Definido pela jurisprudência, a partir da inércia do exequente. "
                         "Não há marco legal único — depende do caso e do tribunal.",
        "observacao": "Aplicar este regime a atos anteriores a 18/03/2016 exige apoio "
                      "jurisprudencial específico; não há dispositivo a citar.",
        "status": "a_revisar",
    },
    {
        "id": "cpc2015_original",
        "rotulo": "CPC/2015, redação original (18/03/2016 a 26/08/2021)",
        "inicio": VIGENCIA_CPC_2015,
        "fim": date(2021, 8, 26),
        "fundamento": "CPC/2015, art. 921, III e §§1º a 5º (redação original)",
        "regra": "Não localizado o devedor ou bens penhoráveis, o juiz suspende a execução "
                 "por 1 ano, durante o qual a prescrição NÃO corre (§1º). Findo esse prazo "
                 "sem bens, os autos vão ao arquivo (§2º). A partir daí corre a prescrição "
                 "intercorrente (§4º), reconhecível de ofício após ouvir as partes (§5º).",
        "termo_inicial": "Fim do ano de suspensão do §1º — ou seja, um ano após a decisão "
                         "que suspendeu a execução.",
        "observacao": "O prazo que corre é o MESMO da pretensão executiva (Súmula 150/STF), "
                      "definido pelo título.",
        "status": "confirmado",
    },
    {
        "id": "lei14195",
        "rotulo": "CPC/2015 com a Lei 14.195/2021 (a partir de 27/08/2021)",
        "inicio": VIGENCIA_LEI_14195,
        "fim": None,
        "fundamento": "CPC/2015, art. 921, §§1º, 4º e 4º-A (redação da Lei 14.195/2021)",
        "regra": "O termo inicial deixa de depender da decisão de suspensão: passa a ser a "
                 "ciência da primeira tentativa infrutífera de localizar o devedor ou bens "
                 "penhoráveis. A suspensão de 1 ano ocorre uma única vez. A efetiva citação, "
                 "intimação ou constrição de bens penhoráveis interfere na contagem.",
        "termo_inicial": "Ciência da PRIMEIRA tentativa infrutífera — em regra anterior à "
                         "decisão de suspensão, o que antecipa o início da contagem.",
        "observacao": "Antecipar o termo inicial pode consumar a prescrição bem antes do que "
                      "o regime anterior indicaria. A aplicação a execuções já em curso em "
                      "27/08/2021 é questão de direito intertemporal em disputa — conferir o "
                      "estado da jurisprudência antes de concluir.",
        "status": "a_revisar",
    },
]


# ── Prazo da pretensão executiva, por título ─────────────────────────────────
# Súmula 150/STF: "Prescreve a execução no mesmo prazo de prescrição da ação."
# É o prazo que corre como intercorrente depois de iniciado o cômputo.

PRAZOS_POR_TITULO = {
    "nota_promissoria": {
        "rotulo": "Nota promissória",
        "prazo": "3 anos",
        "termo": "Do vencimento",
        "fundamento": "LUG (Decreto 57.663/1966), art. 70",
        "status": "confirmado",
    },
    "letra_de_cambio": {
        "rotulo": "Letra de câmbio",
        "prazo": "3 anos",
        "termo": "Do vencimento",
        "fundamento": "LUG (Decreto 57.663/1966), art. 70",
        "status": "confirmado",
    },
    "duplicata": {
        "rotulo": "Duplicata",
        "prazo": "3 anos contra o sacado",
        "termo": "Do vencimento",
        "fundamento": "Lei 5.474/1968, art. 18, I",
        "observacao": "Contra endossante e avalista os prazos são menores (art. 18, II e III).",
        "status": "confirmado",
    },
    "cheque": {
        "rotulo": "Cheque",
        "prazo": "6 meses para a execução",
        "termo": "Do fim do prazo de apresentação",
        "fundamento": "Lei 7.357/1985, art. 59",
        "observacao": "Perdida a força executiva: ação de enriquecimento em 2 anos (art. 61) "
                      "e monitória em 5 anos (Súmula 503/STJ).",
        "status": "confirmado",
    },
    "instrumento_particular": {
        "rotulo": "Contrato / instrumento particular de dívida líquida",
        "prazo": "5 anos",
        "termo": "Do vencimento da obrigação",
        "fundamento": "CC/2002, art. 206, §5º, I",
        "observacao": "Alcança confissão de dívida e contratos assinados por duas testemunhas "
                      "(CPC/2015, art. 784, III).",
        "status": "confirmado",
    },
    "ccb": {
        "rotulo": "Cédula de Crédito Bancário (CCB)",
        "prazo": "DIVERGENTE — 3 anos ou 5 anos",
        "termo": "Do vencimento",
        "fundamento": "Lei 10.931/2004, arts. 26 a 28",
        "divergencia": [
            "3 anos: a CCB é título de crédito e atrai a disciplina cambial (LUG, art. 70), "
            "por força da aplicação subsidiária da legislação cambiária.",
            "5 anos: prevalece o art. 206, §5º, I, do CC/2002, por se tratar de dívida líquida "
            "constante de instrumento particular.",
        ],
        "observacao": "É o título mais comum da carteira e o de prazo mais disputado. Não "
                      "adotar uma das posições sem conferir o entendimento atual do STJ e do "
                      "tribunal de origem: a diferença de 2 anos decide o caso.",
        "status": "a_revisar",
    },
    "cpr": {
        "rotulo": "Cédula de Produto Rural (CPR e CPR-F)",
        "prazo": "A conferir — em regra 3 anos, pela disciplina cambial",
        "termo": "Do vencimento",
        "fundamento": "Lei 8.929/1994 (aplicação subsidiária da legislação cambial)",
        "observacao": "A CPR-F (liquidação financeira) e a CPR física podem receber tratamento "
                      "distinto. Conferir antes de fundamentar.",
        "status": "a_revisar",
    },
    "cedula_rural": {
        "rotulo": "Cédula rural (pignoratícia, hipotecária, CCR)",
        "prazo": "A conferir — em regra 3 anos, pela disciplina cambial",
        "termo": "Do vencimento",
        "fundamento": "Decreto-lei 167/1967",
        "status": "a_revisar",
    },
    "honorarios_advocaticios": {
        "rotulo": "Honorários advocatícios",
        "prazo": "5 anos",
        "termo": "Do trânsito em julgado da decisão que os fixou",
        "fundamento": "CC/2002, art. 206, §5º, II; Lei 8.906/1994, art. 25",
        "status": "confirmado",
    },
    "alugueis": {
        "rotulo": "Aluguéis",
        "prazo": "3 anos",
        "termo": "Do vencimento de cada parcela",
        "fundamento": "CC/2002, art. 206, §3º, I",
        "status": "confirmado",
    },
}


# ── Regra de transição do CC/2002 ────────────────────────────────────────────

TRANSICAO = [
    {
        "id": "art_2028",
        "rotulo": "Prazos a cavaleiro da entrada em vigor do CC/2002",
        "fundamento": "CC/2002, art. 2.028",
        "regra": "Aplica-se o prazo da lei anterior quando (a) o CC/2002 REDUZIU o prazo E "
                 "(b) em 11/01/2003 já havia transcorrido MAIS DA METADE do prazo da lei "
                 "revogada. Faltando qualquer das duas condições, aplica-se o prazo novo, "
                 "contado a partir de 11/01/2003.",
        "status": "confirmado",
    },
    {
        "id": "intertemporal_921",
        "rotulo": "Execuções em curso quando mudou o regime do art. 921",
        "fundamento": "Direito intertemporal processual (isolamento dos atos processuais)",
        "regra": "Cada ato processual rege-se pela lei vigente à sua data. Um mesmo processo "
                 "pode atravessar os três regimes: o enquadramento é feito intervalo a "
                 "intervalo, não pelo regime vigente na distribuição nem pelo atual.",
        "observacao": "A aplicação do termo inicial da Lei 14.195/2021 a fatos anteriores a "
                      "27/08/2021 é controvertida. Conferir o estado da jurisprudência.",
        "status": "a_revisar",
    },
]


# ── Mitigantes: o que impede, suspende, interrompe ou afasta ─────────────────

MITIGANTES = [
    {"efeito": "afasta", "rotulo": "Bens penhoráveis localizados / penhora efetivada",
     "detalhe": "A execução deixa de estar na hipótese do art. 921, III. Sem a causa da "
                "suspensão, não há intercorrente a contar naquele intervalo.",
     "fundamento": "CPC/2015, art. 921, III, a contrario sensu", "status": "confirmado"},
    {"efeito": "interrompe", "rotulo": "Reconhecimento da dívida pelo devedor",
     "detalhe": "Parcelamento, acordo, confissão ou pagamento parcial interrompem a "
                "prescrição e zeram a contagem.",
     "fundamento": "CC/2002, art. 202, VI", "status": "confirmado"},
    {"efeito": "suspende", "rotulo": "Suspensão convencional a pedido do exequente",
     "detalhe": "Suspensão por convenção das partes; não é a hipótese do art. 921, III e "
                "não deflagra a contagem da intercorrente.",
     "fundamento": "CPC/2015, art. 922", "status": "confirmado"},
    {"efeito": "suspende", "rotulo": "Recuperação judicial do devedor",
     "detalhe": "As execuções contra o devedor ficam suspensas; o período não corre contra "
                "o credor.",
     "fundamento": "Lei 11.101/2005, art. 6º", "status": "confirmado"},
    {"efeito": "suspende", "rotulo": "Embargos à execução com efeito suspensivo",
     "detalhe": "Enquanto suspensa a execução por decisão judicial, não há inércia imputável "
                "ao exequente.",
     "fundamento": "CPC/2015, art. 919, §1º", "status": "confirmado"},
    {"efeito": "afasta", "rotulo": "Demora imputável ao Judiciário",
     "detalhe": "A parte não é prejudicada pela demora dos mecanismos do Poder Judiciário. "
                "Súmula editada para a citação, aplicada por analogia à intercorrente.",
     "fundamento": "Súmula 106/STJ", "status": "a_revisar"},
    {"efeito": "afasta", "rotulo": "Diligências efetivas do exequente",
     "detalhe": "Pedidos de Sisbajud, Renajud, Infojud, penhora no rosto dos autos e "
                "desconsideração afastam a inércia — desde que efetivamente requeridos e "
                "apreciados, com data e referência nos autos.",
     "fundamento": "Ausência de inércia (pressuposto da intercorrente)", "status": "a_revisar"},
    {"efeito": "nulidade", "rotulo": "Reconhecimento sem intimação prévia do exequente",
     "detalhe": "O juiz só pode reconhecer de ofício depois de ouvidas as partes. A falta "
                "dessa intimação é vício da decisão que extinguiu a execução — é tese de "
                "ataque quando o crédito já foi declarado prescrito.",
     "fundamento": "CPC/2015, art. 921, §5º", "status": "confirmado"},
    {"efeito": "suspende", "rotulo": "Incidente de desconsideração em curso",
     "detalhe": "Instaurado o IDPJ, a execução fica suspensa quanto ao incidente.",
     "fundamento": "CPC/2015, art. 134, §3º", "status": "a_revisar"},
]


TIPOS_DE_MARCO = [
    "distribuicao", "citacao",
    # Atividade do exequente. Não mexem na contagem (ver _ZERA/_PAUSA/_RETOMA abaixo),
    # mas são a prova documental da tese de "ausência de inércia" — sem elas na
    # cronologia não há como sustentar a diligência do credor data por data.
    "manifestacao_exequente", "pedido_penhora", "pedido_andamento",
    "tentativa_infrutifera", "suspensao_921",
    "arquivamento", "penhora", "bens_localizados", "parcelamento", "embargos",
    "recuperacao_judicial", "retomada", "extincao", "outro",
]


# ── Enquadramento determinístico ─────────────────────────────────────────────

def _para_data(valor) -> date | None:
    """Aceita date ou 'DD/MM/AAAA' (com ou sem texto em volta)."""
    if isinstance(valor, date):
        return valor
    achado = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", str(valor or ""))
    if not achado:
        return None
    try:
        return date(int(achado.group(3)), int(achado.group(2)), int(achado.group(1)))
    except ValueError:
        return None


def regime_aplicavel(valor) -> dict | None:
    """Regime do art. 921 vigente na data informada.

    Determinístico e em Python de propósito: é o tipo de conta que o modelo erra de
    forma silenciosa e que ninguém consegue auditar depois.
    """
    quando = _para_data(valor)
    if quando is None:
        return None
    for regime in REGIMES:
        depois_do_inicio = regime["inicio"] is None or quando >= regime["inicio"]
        antes_do_fim = regime["fim"] is None or quando <= regime["fim"]
        if depois_do_inicio and antes_do_fim:
            return regime
    return None


def marcos_no_intervalo(inicio, fim) -> list[dict]:
    """Marcos de vigência que caem entre duas datas — as viradas de regime que a
    cronologia precisa mostrar entre um ato e o seguinte."""
    de, ate = _para_data(inicio), _para_data(fim)
    if de is None or ate is None:
        return []
    if de > ate:
        de, ate = ate, de
    return [m for m in MARCOS_LEGAIS if de < m["data"] <= ate]


def prazo_do_titulo(descricao) -> dict | None:
    """Casa a descrição livre do lastro com uma entrada de PRAZOS_POR_TITULO."""
    texto = str(descricao or "").casefold()
    if not texto:
        return None
    chaves = [
        ("ccb", ("ccb", "cédula de crédito bancário", "cedula de credito bancario")),
        ("cpr", ("cpr", "cédula de produto rural", "cedula de produto rural")),
        ("cedula_rural", ("cédula rural", "cedula rural", "pignoratícia", "pignoraticia",
                          "cédula hipotecária", "decreto-lei 167")),
        ("nota_promissoria", ("nota promissória", "nota promissoria", "promissória")),
        ("letra_de_cambio", ("letra de câmbio", "letra de cambio")),
        ("duplicata", ("duplicata",)),
        ("cheque", ("cheque",)),
        ("honorarios_advocaticios", ("honorário", "honorario")),
        ("alugueis", ("aluguel", "aluguéis", "alugueis", "locação", "locacao")),
        ("instrumento_particular", ("contrato", "confissão de dívida", "confissao de divida",
                                    "instrumento particular", "abertura de crédito",
                                    "abertura de credito")),
    ]
    for chave, termos in chaves:
        if any(termo in texto for termo in termos):
            return dict(PRAZOS_POR_TITULO[chave], id=chave)
    return None


# ── Bloco para o prompt ──────────────────────────────────────────────────────

def _linha_status(entrada: dict) -> str:
    return " [CONFERIR]" if entrada.get("status") == "a_revisar" else ""


def bloco_prompt() -> str:
    """Renderiza a base como texto para o prompt.

    O modelo recebe as regras prontas em vez de recordá-las: o que se pede a ele é
    aplicar aos fatos do processo, não lembrar de legislação.
    """
    partes = ["═══ BASE DE REGRAS — PRESCRIÇÃO INTERCORRENTE (EXECUÇÃO COMUM) ═══",
              "Use EXCLUSIVAMENTE as regras abaixo. Não recorra a memória própria sobre",
              "prazos ou dispositivos. Itens marcados [CONFERIR] têm divergência ou",
              "dependem de jurisprudência: ao usá-los, diga isso expressamente no texto.",
              "",
              "REGIMES POR DATA DO ATO (o processo pode atravessar mais de um):"]
    for regime in REGIMES:
        partes.append(f"- {regime['rotulo']}{_linha_status(regime)}")
        partes.append(f"    Fundamento: {regime['fundamento']}")
        partes.append(f"    Regra: {regime['regra']}")
        partes.append(f"    Termo inicial: {regime['termo_inicial']}")
        if regime.get("observacao"):
            partes.append(f"    Atenção: {regime['observacao']}")

    partes += ["", "PRAZO DA PRETENSÃO EXECUTIVA POR TÍTULO (Súmula 150/STF — é o prazo que",
               "corre como intercorrente):"]
    for entrada in PRAZOS_POR_TITULO.values():
        partes.append(f"- {entrada['rotulo']}: {entrada['prazo']} ({entrada['termo']})"
                      f" — {entrada['fundamento']}{_linha_status(entrada)}")
        for posicao in entrada.get("divergencia", []):
            partes.append(f"    · {posicao}")
        if entrada.get("observacao"):
            partes.append(f"    Atenção: {entrada['observacao']}")

    partes += ["", "DIREITO INTERTEMPORAL:"]
    for entrada in TRANSICAO:
        partes.append(f"- {entrada['rotulo']}{_linha_status(entrada)}: {entrada['regra']}"
                      f" ({entrada['fundamento']})")
        if entrada.get("observacao"):
            partes.append(f"    Atenção: {entrada['observacao']}")

    partes += ["", "MITIGANTES — verificar TODOS antes de afirmar risco:"]
    for entrada in MITIGANTES:
        partes.append(f"- [{entrada['efeito'].upper()}] {entrada['rotulo']}"
                      f"{_linha_status(entrada)}: {entrada['detalhe']} ({entrada['fundamento']})")

    partes += [
        "",
        "COMO CONCLUIR:",
        "- O risco é apresentado como RISCO, com o fundamento e a referência processual",
        "  ao lado. Nunca afirme prescrição consumada como fato.",
        "- Se faltar data para fechar a conta (ex.: não consta a ciência da tentativa",
        "  infrutífera), diga qual data falta e onde ela seria encontrada.",
        "- Aponte cada mitigante encontrado nos autos, com data e referência.",
        "",
    ]
    return "\n".join(partes)


# ══════════════════════════════════════════════════════════════════════════════
# Cálculo
#
# A base acima é o que a lei diz. Aqui a lei vira conta: quando a prescrição começou a
# correr, quanto correu de fato e se o lapso já superou o prazo. Determinístico, em
# Python — o modelo só entrega fatos datados.
#
# É um MODELO de contagem, com escolhas declaradas abaixo. Mostra a conta e onde ela
# pode virar; não substitui a leitura do analista.
# ══════════════════════════════════════════════════════════════════════════════

# Prazo da pretensão executiva, em anos. A CCB traz as duas posições em disputa: a
# diferença de 2 anos é o que decide o caso, então as duas contas são feitas.
ANOS_POR_TITULO = {
    "nota_promissoria": [3], "letra_de_cambio": [3], "duplicata": [3],
    "cedula_rural": [3], "cpr": [3], "alugueis": [3],
    "cheque": [0.5],
    "instrumento_particular": [5], "honorarios_advocaticios": [5],
    "ccb": [3, 5],
}

# CC/1916, art. 177: as ações pessoais prescreviam em 20 anos.
ANOS_CC_1916 = 20

# Duração da suspensão do art. 921 quando o juiz não fixa prazo — é o que o IAC 1/STJ
# (REsp 1.604.412/SC) assentou para a execução comum sob o CPC/1973.
MESES_SUSPENSAO_PADRAO = 12

_RE_DURACAO = re.compile(r"(\d+)\s*(anos?|m[eê]s|meses|dias?)", re.IGNORECASE)


def meses_da_duracao(texto):
    """"1 ano" -> 12, "6 meses" -> 6, "90 dias" -> 3. Sem duração legível, None."""
    achado = _RE_DURACAO.search(str(texto or ""))
    if not achado:
        return None
    quantidade, unidade = int(achado.group(1)), achado.group(2).lower()
    if unidade.startswith("ano"):
        return quantidade * 12
    if unidade.startswith("dia"):
        return max(1, round(quantidade / 30))
    return quantidade


_DIAS_NO_MES = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]


def _somar_meses(quando: date, meses: int) -> date:
    ano = quando.year + (quando.month - 1 + meses) // 12
    mes = (quando.month - 1 + meses) % 12 + 1
    limite = _DIAS_NO_MES[mes - 1]
    if mes == 2 and (ano % 4 == 0 and (ano % 100 != 0 or ano % 400 == 0)):
        limite = 29
    return date(ano, mes, min(quando.day, limite))


def _somar_anos(quando: date, anos: float) -> date:
    return _somar_meses(quando, round(anos * 12))


def prazo_com_transicao(vencimento, chave_titulo=None) -> list:
    """Prazo aplicável ao título, já resolvida a transição do art. 2.028 do CC/2002.

    Antes de 11/01/2003 a dívida particular prescrevia em 20 anos (CC/1916, art. 177).
    O art. 2.028 manda manter o prazo antigo só quando o novo o REDUZIU e já havia
    corrido MAIS DA METADE do prazo velho na entrada em vigor. Fora disso vale o prazo
    novo, contado de 11/01/2003 — e é isso que muda o resultado em execução antiga.
    """
    anos_novos = ANOS_POR_TITULO.get(chave_titulo or "", [5])
    inicio = _para_data(vencimento)
    if inicio is None or inicio >= VIGENCIA_CC_2002:
        return [{"anos": a, "conta_de": inicio,
                 "regra": "Prazo do CC/2002, contado do vencimento do título.",
                 "status": "confirmado"} for a in anos_novos]

    corridos = (VIGENCIA_CC_2002 - inicio).days / 365.25
    if corridos > ANOS_CC_1916 / 2:
        return [{
            "anos": ANOS_CC_1916, "conta_de": inicio,
            "regra": (f"Vale o prazo do CC/1916 (20 anos): em 11/01/2003 já haviam corrido "
                      f"{corridos:.1f} anos, mais da metade dos 20 (CC/2002, art. 2.028)."),
            "status": "confirmado",
        }]
    return [{
        "anos": a, "conta_de": VIGENCIA_CC_2002,
        "regra": (f"Vale o prazo do CC/2002: em 11/01/2003 haviam corrido só {corridos:.1f} "
                  f"anos dos 20 do CC/1916 — menos da metade —, então o prazo novo corre "
                  f"a partir de 11/01/2003 (art. 2.028)."),
        "status": "confirmado",
    } for a in anos_novos]


def termo_inicial_intercorrente(marcos: list):
    """Quando a intercorrente começou a correr, pelo regime vigente no ato.

    - Lei 14.195/2021: da ciência da primeira tentativa infrutífera.
    - CPC/2015 original: do fim do ano de suspensão do art. 921, §1º.
    - CPC/1973 (IAC 1/STJ): do fim do prazo de suspensão FIXADO pelo juiz; não havendo
      prazo fixado, um ano depois da decisão que suspendeu.
    """
    def _primeiro(tipo):
        return next((m for m in marcos
                     if m.get("tipo") == tipo and _para_data(m.get("data"))), None)

    suspensao, tentativa = _primeiro("suspensao_921"), _primeiro("tentativa_infrutifera")

    if suspensao:
        quando = _para_data(suspensao["data"])
        regime = regime_aplicavel(quando)
        if regime and regime["id"] == "lei14195" and tentativa:
            return {"data": _para_data(tentativa["data"]), "marco": tentativa, "regime": regime,
                    "regra": "Lei 14.195/2021: conta da ciência da primeira tentativa "
                             "infrutífera de localizar o devedor ou bens penhoráveis."}
        fixados = meses_da_duracao(suspensao.get("duracao"))
        meses = fixados or MESES_SUSPENSAO_PADRAO
        if regime and regime["id"] == "cpc1973":
            regra = (f"IAC 1/STJ: conta do fim do prazo de suspensão fixado pelo juiz "
                     f"({meses} meses)." if fixados else
                     "IAC 1/STJ: não houve prazo de suspensão fixado nos autos, então "
                     "conta um ano depois da decisão que suspendeu.")
        else:
            regra = f"Art. 921, §1º: conta do fim do ano de suspensão ({meses} meses)."
        return {"data": _somar_meses(quando, meses), "marco": suspensao,
                "regime": regime, "regra": regra}

    if tentativa:
        quando = _para_data(tentativa["data"])
        regime = regime_aplicavel(quando)
        if regime and regime["id"] == "lei14195":
            return {"data": quando, "marco": tentativa, "regime": regime,
                    "regra": "Lei 14.195/2021: conta da ciência da primeira tentativa "
                             "infrutífera, independentemente de decisão de suspensão."}
        return {"data": _somar_meses(quando, MESES_SUSPENSAO_PADRAO), "marco": tentativa,
                "regime": regime,
                "regra": "Não há decisão de suspensão nos autos; contado um ano da ciência "
                         "da não localização, por analogia ao art. 921, §1º. Conferir."}
    return None


# Efeito de cada marco sobre a contagem, a partir do termo inicial.
#
# Pedido de penhora, pedido de andamento e manifestação do exequente ficam FORA dos três
# conjuntos de propósito: requerer não retira a execução da hipótese do art. 921, III —
# quem retira é o resultado (penhora efetivada, bens localizados, parcelamento). Tratá-los
# como se zerassem faria o robô devolver "não prescreveu" para toda execução em que o
# credor peticionou de tempo em tempo sem nunca achar nada, que é justamente o caso
# típico. Eles entram na cronologia como prova da diligência do credor — matéria da tese
# de ausência de inércia (ver MARCOS_LEGAIS), avaliada no parecer, não na conta.
_ZERA = {"penhora", "bens_localizados", "parcelamento", "retomada", "citacao"}
_PAUSA = {"recuperacao_judicial", "embargos"}
_RETOMA = {"retomada", "tentativa_infrutifera", "suspensao_921"}


def lapsos_de_inercia(marcos: list, inicio: date, fim: date) -> list:
    """Quebra o período em trechos e diz, de cada um, se a prescrição corre.

    Modelo de contagem, declarado: penhora efetivada, localização de bens, parcelamento
    e retomada ZERAM o acumulado — a execução deixa de estar na hipótese do art. 921,
    III, que é o pressuposto da intercorrente. Recuperação judicial e embargos PAUSAM
    sem zerar. Nos demais trechos, corre.
    """
    pontos = [(inicio, None)]
    for marco in marcos:
        quando = _para_data(marco.get("data"))
        if quando and inicio < quando <= fim:
            pontos.append((quando, marco))
    pontos.append((fim, None))
    pontos.sort(key=lambda p: p[0])

    trechos, pausado, acumulado = [], False, 0.0
    for (de, _), (ate, marco_fim) in zip(pontos, pontos[1:]):
        if ate > de:
            anos = (ate - de).days / 365.25
            if not pausado:
                acumulado += anos
            trechos.append({"de": de, "ate": ate, "anos": anos, "corre": not pausado,
                            "acumulado": acumulado,
                            "motivo": "" if not pausado else "suspenso"})
        tipo = (marco_fim or {}).get("tipo")
        if tipo in _ZERA:
            acumulado, pausado = 0.0, False
            trechos.append({"de": ate, "ate": ate, "anos": 0.0, "corre": False,
                            "acumulado": 0.0, "motivo": f"zera a contagem ({tipo})"})
        elif tipo in _PAUSA:
            pausado = True
        elif tipo in _RETOMA:
            pausado = False
    return trechos


def avaliar(marcos: list, titulo=None, vencimento=None, hoje=None) -> dict:
    """Termo inicial, lapsos e veredito — um cenário por prazo em disputa."""
    hoje = hoje or date.today()
    marcos = [m for m in (marcos or []) if isinstance(m, dict)]
    entrada = prazo_do_titulo(titulo)
    chave = (entrada or {}).get("id")

    termo = termo_inicial_intercorrente(marcos)
    extincao = next((m for m in marcos
                     if m.get("tipo") == "extincao" and _para_data(m.get("data"))), None)
    fim = _para_data(extincao["data"]) if extincao else hoje

    if vencimento:
        prazos = prazo_com_transicao(vencimento, chave)
    else:
        prazos = [{"anos": a, "conta_de": None, "status": "a_revisar",
                   "regra": "Sem data de vencimento nos autos, a transição do art. 2.028 "
                            "não pôde ser conferida."}
                  for a in ANOS_POR_TITULO.get(chave or "", [5])]

    cenarios = []
    for prazo in prazos:
        item = dict(prazo)
        if termo and termo["data"] and fim > termo["data"]:
            trechos = lapsos_de_inercia(marcos, termo["data"], fim)
            corrido = trechos[-1]["acumulado"] if trechos else 0.0
            item.update({"trechos": trechos, "corrido": corrido,
                         "consumado": corrido >= prazo["anos"],
                         "faltam": max(0.0, prazo["anos"] - corrido),
                         "data_limite": _somar_anos(termo["data"], prazo["anos"])})
        else:
            item.update({"trechos": [], "corrido": 0.0, "consumado": False,
                         "faltam": prazo["anos"], "data_limite": None})
        cenarios.append(item)

    return {"termo": termo, "fim": fim, "extincao": extincao, "titulo": entrada,
            "cenarios": cenarios, "divergente": len(cenarios) > 1}
