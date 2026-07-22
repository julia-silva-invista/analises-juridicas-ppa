# -*- coding: utf-8 -*-
import os
import re
import math
import unicodedata
import time
import tempfile
import concurrent.futures
from datetime import date
from pathlib import Path
from typing import Optional, List

from dateutil.relativedelta import relativedelta

import pandas as pd
from pydantic import BaseModel
from google import genai
from google.genai import types

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from utils import _retry, _responder_pergunta_generica


def _mat_get_clients():
    k1 = os.getenv("GEMINI_API_KEY_1") or os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
    k2 = os.getenv("GEMINI_API_KEY_2") or k1
    if not k1:
        raise RuntimeError("GEMINI_API_KEY_1 não configurada.")
    return genai.Client(api_key=k1), genai.Client(api_key=k2)


# ── Modelos Pydantic ──────────────────────────────────────────────────────────

class OnusFinanceiro(BaseModel):
    codigo: Optional[str] = None
    tipo: Optional[str] = None          # penhora, hipoteca, alienacao fiduciaria, CCB, etc.
    numero_processo: Optional[str] = None   # nº do processo judicial (penhora, arresto, execucao)
    numero_instrumento: Optional[str] = None  # nº da CCB, contrato, operacao bancaria
    valor_principal: Optional[float] = None
    data_celebracao: Optional[str] = None
    parcela_mensal: Optional[float] = None
    vencimento_final: Optional[str] = None
    cancelado: Optional[bool] = False   # True se o ato foi baixado/cancelado por AV/R posterior


class Transmissao(BaseModel):
    codigo: Optional[str] = None        # ex: "R.4"
    tipo: Optional[str] = None          # Compra e Venda, Doação, Integralização, etc.
    data: Optional[str] = None          # DD/MM/AAAA
    de_nome: Optional[str] = None
    para_nome: Optional[str] = None
    de_doc: Optional[str] = None        # CPF/CNPJ sem formatacao (so digitos)
    para_doc: Optional[str] = None


class MatriculaExtraida(BaseModel):
    matricula: Optional[str] = None
    comarca: Optional[str] = None
    proprietario_atual: Optional[str] = None
    descricao_imovel: Optional[str] = None
    transmissoes_averbadas_registradas: Optional[str] = None
    onus_vigentes_registrados_averbados: Optional[str] = None
    observacoes: Optional[str] = None
    onus_cancelados: Optional[str] = None
    grau_confianca: Optional[str] = None
    onus_financeiros: Optional[List[OnusFinanceiro]] = None
    transmissoes_estruturadas: Optional[List[Transmissao]] = None


# ── Prompt ────────────────────────────────────────────────────────────────────

PROMPT_MATRICULA = (
    "Voce e um assistente especializado em analise de matriculas de imoveis brasileiros.\n\n"
    "Analise integralmente o PDF enviado e extraia as informacoes em formato estruturado.\n\n"
    "Preencha os seguintes campos:\n"
    "- matricula\n- comarca\n- proprietario_atual\n- descricao_imovel\n"
    "- transmissoes_averbadas_registradas\n- onus_vigentes_registrados_averbados\n"
    "- observacoes\n- onus_cancelados\n- grau_confianca\n"
    "- onus_financeiros\n- transmissoes_estruturadas\n\n"
    "Regras obrigatorias:\n\n"
    "1. Proprietario atual: ultimo registro translativo valido. APENAS nomes, sem CPF/CNPJ.\n"
    "   Se a matrícula estiver encerrada, indique o proprietario e registre em observacoes.\n"
    "   Se pertencer a mais de uma pessoa, indique o percentual de cada titular.\n"
    "   Exemplo: Fulano da Silva (40%) Beltrano Camargo (60%)\n\n"
    "2. comarca: indicar a comarca e qual é o cartório. Exemplo: '1º CRI de Curitiba/PR'.\n\n"
    "3. descricao_imovel: resumo em ate 3 frases. Tipo, localizacao, area, identificadores.\n\n"
    "4. transmissoes_averbadas_registradas: todas, cronologicas, sinteticas.\n"
    "   Inclua o CPF/CNPJ de cada pessoa, mas APENAS NA PRIMEIRA VEZ que ela aparecer nesta celula.\n"
    "   Diga sempre o preço do imóvel em caso de compra e venda ou integralizacao.\n"
    "   Inclua registros anteriores a abertura se disponiveis.\n"
    "   Exemplo:\n"
    "   R.4 - **Compra e Venda**: Jose Marques da Cruz (CPF 123.456.789-00) vendeu para Carlos Guimaraes"
    " (CPF 987.654.321-00) em 15/12/1986, pelo valor de R$ 2.000.000,00.\n\n"
    "   R.5 - **Doação**: Carlos Guimaraes doou para Murilo Menezes (CPF 111.222.333-44) em 03/06/1988.\n\n"
    "   Sempre que trouxer a natureza das transmissões, coloque em negrito apenas o tipo (ex: **Compra e Venda**).\n\n"
    "5. onus_vigentes_registrados_averbados: apenas gravames vigentes de natureza financeira.\n"
    "   NAO inclua aqui: indisponibilidade de bens, averbacao premonitoria, averbacao de ajuizamento,"
    " averbacao de execucao, protesto, restricao administrativa. Estes vao em observacoes.\n"
    "   Se nao houver onus vigentes financeiros, preencha exatamente: Sem onus vigentes identificados.\n\n"
    "6. Formato de cada onus vigente:\n"
    "   [Codigo]: [Tipo], [n. cedula se disponivel], [dd/mm/aaaa], [partes — sem CPF/CNPJ],\n"
    "   vencimento [data se disponivel], Valor principal: R$ [valor se disponivel].\n"
    "   OBRIGATORIO para penhora, arresto, execucao ou ajuizamento: inclua o numero do processo.\n"
    "   Exemplo: AV.12: Penhora, 12/03/2023, Exequente: Banco XYZ; Executado: Joao Silva,"
    " Proc. 1234567-89.2023.8.26.0001, Valor principal: R$ 250.000,00.\n"
    "   Vara e comarca vao em observacoes, nao aqui.\n\n"
    "7. Nao inclua atos cancelados em onus_vigentes. Se cancelado, mencione em onus_cancelados.\n\n"
    "8. PDF escaneado: extraia o maximo possivel do conteudo visual.\n\n"
    "9. Nunca invente dados. Nomes: primeira letra maiuscula, exceto artigos (de, da, do, das, dos, e).\n"
    "   Erros de grafia: copie como esta e registre em observacoes.\n\n"
    "10. observacoes: inclua SEMPRE que houver:\n"
    "   - Indisponibilidade de bens (com n. processo se disponivel)\n"
    "   - Averbacao premonitoria (com n. processo se disponivel)\n"
    "   - Averbacao de ajuizamento de execucao (com n. processo se disponivel)\n"
    "   - Averbacoes relevantes (divorcio, matrimonio, georreferenciamento, reserva legal, quitacao)\n"
    "   - Irregularidades, CPFs divergentes para a mesma pessoa, erros de grafia\n"
    "   - Matrícula encerrada, rematricula, baixa legibilidade\n"
    "   NAO inclua onus cancelados em observacoes — eles vao EXCLUSIVAMENTE em onus_cancelados.\n"
    "   NAO inclua vara/comarca/n. processo de penhoras em observacoes — essas infos vao no campo onus_vigentes.\n"
    "   Se nao houver: Sem observacoes relevantes. Nunca deixe vazio.\n\n"
    "11. onus_cancelados: APENAS onus cancelados, sempre indicando qual ato cancelou.\n"
    "    Exemplo: AV.9: Penhora cancelada pela AV.22.\n"
    "    NAO repita esses onus em observacoes.\n"
    "    Se nao houver: Sem onus cancelados identificados.\n\n"
    "12. grau_confianca: Alto, Medio ou Baixo.\n"
    "    Medio/Baixo: justificativa entre parenteses. Ex: Medio (Baixa legibilidade da AV-3).\n"
    "    Alto: preencha apenas Alto.\n"
    "    A marca d'agua 'nao possui valor de certidao' por si so nao e indicio de confianca reduzida.\n"
    "    ATENCAO: matricula encerrada ou rematriculada = obrigatoriamente Medio ou Baixo.\n\n"
    "13. Responda apenas com JSON valido aderente ao schema.\n\n"
    "14. Redacao padronizada e objetiva. Evite textos longos.\n\n"
    "15. Pule UMA LINHA entre itens autonomos diferentes em cada campo de texto.\n\n"
    "16. NUNCA pule linha dentro de uma mesma frase. Intervalos como AV.1 a AV.8 ou"
    " referencias como conforme R.5 sao parte da frase e NAO devem ter quebra de linha.\n\n"
    "17. Apos a analise, verifique todos os AVs/Rs para ver se algum ficou faltando.\n"
    "    Caso falte algum, indique em observacoes qual registro ou averbaçao nao pode ser identificado.\n\n"
    "18. Para cada onus de natureza financeira encontrado (penhora, hipoteca, alienacao fiduciaria,"
    " CCB, confissao de divida ou qualquer titulo de credito), preencha onus_financeiros com:\n"
    "   - codigo: codigo do ato (ex: R.5, AV.9)\n"
    "   - tipo: tipo do onus (ex: penhora, hipoteca, alienacao fiduciaria, CCB)\n"
    "   - numero_processo: OBRIGATORIO para penhora, arresto, execucao e qualquer ato judicial."
    " Informe o numero completo do processo (ex: 0001234-56.2023.8.26.0100). Null se nao for ato judicial.\n"
    "   - numero_instrumento: OBRIGATORIO para CCB, contrato bancario, operacao de credito, cedula."
    " Informe o numero da CCB, do contrato ou da operacao (ex: CCB 12345, Contrato 67890, Op. 00123)."
    " Null se nao for instrumento de credito.\n"
    "   - valor_principal: valor em reais (numero puro, sem R$ ou pontos)\n"
    "   - data_celebracao: data do ato em DD/MM/AAAA\n"
    "   - parcela_mensal: valor da parcela mensal se explicitamente indicada. Null se nao houver.\n"
    "   - vencimento_final: data de vencimento final em DD/MM/AAAA. Null se nao houver.\n"
    "   - cancelado: true se houver AV/R posterior que baixou, cancelou ou levantou este onus."
    " false se ainda vigente.\n"
    "   NAO inclua: indisponibilidade, averbacao premonitoria, averbacao de ajuizamento,"
    " averbacao de execucao, protesto.\n"
    "   Averbacao de execucao e averbacao premonitoria sao NOTACOES PROCESSUAIS, nao gravames financeiros"
    " — ignorar no calculo mesmo que tenham valor indicado.\n"
    "   NAO inclua onus sem valor principal identificavel. Nao invente dados.\n\n"
    "19. VARREDURA OBRIGATORIA AV/R POR AV/R:\n"
    "    a) Identifique a numeracao maxima de R e de AV no documento.\n"
    "    b) Percorra individualmente de R.1 ate o ultimo R, e de AV.1 ate o ultimo AV.\n"
    "    c) Para CADA item, classifique em EXATAMENTE uma destas categorias:\n"
    "       - ONUS VIGENTE FINANCEIRO: penhora ativa, hipoteca, alienacao fiduciaria, CCB, arresto"
    " vigente, confissao de divida, titulo de credito nao cancelado -> onus_vigentes\n"
    "       - ONUS NAO FINANCEIRO (vai em observacoes, NAO em onus_vigentes nem em onus_financeiros):"
    " indisponibilidade de bens, averbacao premonitoria, averbacao de ajuizamento, averbacao de execucao, protesto\n"
    "       - ONUS CANCELADO: onus cuja AV/R posterior cancelou/baixou -> onus_cancelados\n"
    "       - OBSERVACAO/TRANSMISSAO: compra e venda, doacao, partilha, integralizacao, divorcio,"
    " retificacao, georreferenciamento, encerramento, etc.\n"
    "    d) NUNCA pule nenhum R/AV. Se ilegivel, cite em observacoes. Omissao e ERRO GRAVE.\n"
    "    e) Antes de finalizar, confira se TODOS os R/AV aparecem em alguma categoria.\n\n"
    "20. transmissoes_estruturadas: para CADA transmissao (compra e venda, doacao, partilha,"
    " integralizacao, etc.) identificada, preencha uma entrada com:\n"
    "   - codigo: codigo do ato (ex: 'R.4')\n"
    "   - tipo: tipo da transmissao (ex: 'Compra e Venda', 'Doacao', 'Integralizacao')\n"
    "   - data: data do ato em DD/MM/AAAA. Null se nao identificada.\n"
    "   - de_nome: nome do transmitente (quem vendeu/doou/transferiu)\n"
    "   - para_nome: nome do adquirente (quem comprou/recebeu)\n"
    "   - de_doc: CPF ou CNPJ do transmitente, APENAS DIGITOS sem formatacao (ex: '12345678901')."
    " Null se nao disponivel.\n"
    "   - para_doc: CPF ou CNPJ do adquirente, APENAS DIGITOS. Null se nao disponivel.\n"
    "   Nao invente. Deixe null qualquer campo nao identificavel.\n\n"
    "21. CONSISTENCIA OBRIGATORIA entre onus_financeiros e onus_vigentes_registrados_averbados:\n"
    "   Antes de finalizar, verifique: para cada item em onus_financeiros com cancelado=false,\n"
    "   ele DEVE estar descrito em onus_vigentes_registrados_averbados.\n"
    "   Se voce nao o incluiu em onus_vigentes, defina cancelado=true para ele.\n"
    "   A inconsistencia (ativo no JSON mas ausente no texto, ou vice-versa) e ERRO GRAVE.\n\n"
    "22. VALORES EM MOEDA PRE-REAL (anterior a julho/1994): qualquer onus cujo valor esteja expresso\n"
    "   em Cruzeiros, Cruzados, NCz$, Cr$, URV ou qualquer moeda anterior ao Real NAO deve ser\n"
    "   incluido em onus_financeiros. Registre-o APENAS em observacoes, indicando a moeda original.\n"
    "   Se nao houver informacao sobre a moeda mas o valor parecer incompativel com imóveis brasileiros\n"
    "   (ex: valores acima de R$ 500.000.000 para credito rural dos anos 80-90), trate como moeda antiga.\n"
)


# ── Helpers de texto ──────────────────────────────────────────────────────────

MINUSC_SET = {"de", "da", "do", "das", "dos", "e", "a", "o", "em", "com", "por", "ao", "aos", "as"}
_ITEM_RE = re.compile(
    r"(?<!\n\n)(?<!\n)(?=(?:R|AV)-\d+[-–]\d[\d.]*\s*(?:\([^)]*\)\s*)?:)",
    re.IGNORECASE,
)


def _mat_capitalizar_token(p):
    core = p.strip(".,;:()")
    suffix = p[len(core):]
    if not core: return p
    low = core.lower()
    if low in MINUSC_SET: return low + suffix
    if "/" in core: return core.upper() + suffix
    if re.match(r"^[A-Z]{2,6}$", core): return core + suffix
    return core.capitalize() + suffix


def _mat_corrigir_caps(texto):
    if not texto or pd.isna(texto): return texto
    resultado = []
    for linha in str(texto).split("\n"):
        nova = []
        for p in linha.split(" "):
            core = p.strip(".,;:()")
            if core.isupper() and len(core) > 2 and "/" not in core and not re.match(r"^\d", core):
                nova.append(_mat_capitalizar_token(p))
            else:
                nova.append(p)
        resultado.append(" ".join(nova))
    return "\n".join(resultado)


def _mat_remover_cpf(texto):
    if not texto or pd.isna(texto): return texto
    t = str(texto)
    t = re.sub(r"\s*\((?:CPF(?:/MF)?|CNPJ)[^)]*\)", "", t, flags=re.IGNORECASE)
    t = re.sub(r",?\s*(?:CPF(?:/MF)?|CNPJ)\s*(?:n[o]?\s*\.?)?\s*[\d]{3}[\d.\-/]+", "", t, flags=re.IGNORECASE)
    t = re.sub(r"\(\s*\)", "", t)
    t = re.sub(r"  +", " ", t)
    t = re.sub(r"\s+([,;.])", r"\1", t)
    return t.strip()


def _mat_normalizar_quebras(texto):
    if not texto or pd.isna(texto): return texto
    t = str(texto).replace("\r\n", "\n").replace("\r", "\n")
    t = _ITEM_RE.sub("\n\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def _mat_normalizar_siglas(texto):
    if not texto or pd.isna(texto): return texto
    t = str(texto)
    t = re.sub(r"\bS/[Aa]\b", "S/A", t)
    t = re.sub(r"\bccir\b", "CCIR", t, flags=re.IGNORECASE)
    return t


def _fmt_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ── Pós-processamento ─────────────────────────────────────────────────────────

def _mat_pos_processar(resultado):
    # Campos onde CPF/CNPJ deve ser removido
    campos_sem_cpf = ["proprietario_atual", "onus_vigentes_registrados_averbados",
                      "observacoes", "onus_cancelados"]
    for campo in campos_sem_cpf:
        v = resultado.get(campo)
        if v:
            resultado[campo] = _mat_remover_cpf(v)

    # Todos os campos de texto: caps e siglas
    campos_texto = ["proprietario_atual", "descricao_imovel",
                    "transmissoes_averbadas_registradas",
                    "onus_vigentes_registrados_averbados", "observacoes", "onus_cancelados"]
    for campo in campos_texto:
        v = resultado.get(campo)
        if v:
            v = _mat_corrigir_caps(v)
            v = _mat_normalizar_siglas(v)
            resultado[campo] = v

    # Normalizar quebras de linha nos campos longos
    for campo in ["transmissoes_averbadas_registradas", "onus_vigentes_registrados_averbados",
                  "observacoes", "onus_cancelados"]:
        v = resultado.get(campo)
        if v:
            resultado[campo] = _mat_normalizar_quebras(_mat_normalizar_siglas(v))

    return resultado


# ── Cálculo de ônus ───────────────────────────────────────────────────────────

_TIPOS_NAO_FINANCEIROS = {
    "indisponibilidade", "premonitoria", "premonitório", "premonotoria",
    "ajuizamento", "protesto", "restricao administrativa",
    "averbacao de execucao", "averbação de execução",
    "averbacao execucao", "averbação execução",
}


def _mat_calcular_valor_onus(onus_list: list, onus_vigentes_texto: str = "",
                              data_referencia: date = None) -> str:
    if not onus_list:
        return ""
    # Se o campo textual já diz "sem ônus vigentes", o modelo concluiu que
    # não há gravames ativos — não computar valor (evita valores em moeda antiga
    # ou itens que o modelo identificou como sem efeito prático serem somados).
    if onus_vigentes_texto and "sem onus vigentes" in onus_vigentes_texto.lower():
        return ""
    if data_referencia is None:
        data_referencia = date.today()

    linhas_individuais = []
    total = 0.0

    for onus in onus_list:
        def _get(key):
            return onus.get(key) if isinstance(onus, dict) else getattr(onus, key, None)

        cancelado = _get("cancelado")
        if cancelado:
            continue

        codigo             = _get("codigo")
        tipo               = _get("tipo") or ""
        numero_processo    = _get("numero_processo")
        numero_instrumento = _get("numero_instrumento")
        valor              = _get("valor_principal")
        data_str           = _get("data_celebracao")
        parcela            = _get("parcela_mensal")

        # Ignora tipos não financeiros (indisponibilidade, premonitória, etc.)
        tipo_lower = tipo.lower()
        if any(t in tipo_lower for t in _TIPOS_NAO_FINANCEIROS):
            continue

        if not valor or not data_str:
            continue
        try:
            partes = data_str.strip().split("/")
            data_onus = date(int(partes[2]), int(partes[1]), int(partes[0]))
        except Exception:
            continue

        delta = relativedelta(data_referencia, data_onus)
        meses = delta.years * 12 + delta.months
        if meses < 0:
            continue

        valor_atualizado = valor * (1 + 0.01 * meses)
        total += valor_atualizado

        prefixo = f"∘ {codigo}" if codigo else "∘"
        if tipo:
            prefixo += f" ({tipo})"
        prefixo += ": "

        linha = prefixo + _fmt_brl(valor_atualizado)

        # Número do processo (penhoras, arrestos, execuções)
        if numero_processo:
            linha += f"\n   Proc. {numero_processo}"

        # Número do instrumento (CCB, contrato, operação)
        if numero_instrumento:
            linha += f"\n   {numero_instrumento}"

        if parcela and parcela > 0:
            saldo = valor_atualizado - parcela * meses
            if saldo > 0:
                linha += f"\n   (em caso de adimplência, saldo de {_fmt_brl(saldo)})"

        linhas_individuais.append(linha)

    if not linhas_individuais:
        return ""

    return f"Total: {_fmt_brl(total)}\n" + "\n".join(linhas_individuais)


# ── Parsing de devedores/relacionados ─────────────────────────────────────────

def _mat_parse_parties(texto: str) -> set:
    """
    Parseia texto com entradas 'CPF/CNPJ — Nome' (uma por linha).
    Retorna set de documentos normalizados (apenas dígitos).
    """
    docs = set()
    if not texto or not texto.strip():
        return docs
    for linha in texto.strip().split("\n"):
        linha = linha.strip()
        if not linha:
            continue
        # Extrai qualquer sequência que pareça CPF ou CNPJ
        for m in re.finditer(r"[\d]{3}[\d.\-/]{3,}[\d]", linha):
            doc_norm = re.sub(r"[^\d]", "", m.group())
            if len(doc_norm) in (11, 14):  # CPF ou CNPJ
                docs.add(doc_norm)
    return docs


def _mat_detectar_alertas(resultado: dict, devedores_docs: set,
                           relacionados_docs: set, data_ajuizamento: date | None) -> set:
    """
    Retorna set com alertas detectados: 'amarelo' e/ou 'vermelho'.
    - amarelo: transmissão envolve devedor ou pessoa do grupo econômico
    - vermelho: transmissão após data de ajuizamento
    """
    alertas = set()
    todos_docs = devedores_docs | relacionados_docs
    transmissoes = resultado.get("transmissoes_estruturadas") or []

    for t in transmissoes:
        if isinstance(t, dict):
            de_doc   = re.sub(r"[^\d]", "", t.get("de_doc") or "")
            para_doc = re.sub(r"[^\d]", "", t.get("para_doc") or "")
            data_str = t.get("data") or ""
        else:
            de_doc   = re.sub(r"[^\d]", "", getattr(t, "de_doc", "") or "")
            para_doc = re.sub(r"[^\d]", "", getattr(t, "para_doc", "") or "")
            data_str = getattr(t, "data", "") or ""

        if todos_docs and (de_doc in todos_docs or para_doc in todos_docs):
            alertas.add("amarelo")

        if data_ajuizamento and data_str:
            try:
                p = data_str.strip().split("/")
                dt = date(int(p[2]), int(p[1]), int(p[0]))
                if dt >= data_ajuizamento:
                    alertas.add("vermelho")
            except Exception:
                pass

    return alertas


# ── Análise de PDF ────────────────────────────────────────────────────────────

def _mat_analisar_pdf(caminho_pdf: str, client) -> dict:
    try:
        caminho_pdf.encode("ascii")
        upload_path = caminho_pdf
    except UnicodeEncodeError:
        import shutil as _shutil
        tmp_ascii = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_ascii.close()
        _shutil.copy2(caminho_pdf, tmp_ascii.name)
        upload_path = tmp_ascii.name

    arquivo = client.files.upload(file=upload_path)

    # Aguarda processamento do arquivo
    total_wait, wait_time = 0, 1
    while total_wait < 120:
        state_name = getattr(getattr(arquivo, "state", None), "name", "")
        if state_name in ("ACTIVE", "FAILED"):
            break
        time.sleep(wait_time)
        total_wait += wait_time
        arquivo = client.files.get(name=arquivo.name)
        wait_time = min(wait_time + 1, 4)

    cfg = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=MatriculaExtraida,
        temperature=0,
        thinking_config=types.ThinkingConfig(thinking_budget=0),
    )

    def _call():
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[PROMPT_MATRICULA, arquivo],
            config=cfg,
        )
        return MatriculaExtraida.model_validate_json(response.text).model_dump()

    resultado = _retry(_call, tentativas=3, espera_base=10)

    try:
        client.files.delete(name=arquivo.name)
    except Exception:
        pass

    return _mat_pos_processar(resultado)


def _mat_worker(args) -> tuple:
    idx, pdf_path, client = args
    try:
        linha = _mat_analisar_pdf(pdf_path, client)
        return idx, linha, None
    except Exception as e:
        return idx, None, str(e)


# ── Helpers de nome ───────────────────────────────────────────────────────────

def _mat_limpar_nome(nome):
    nome = unicodedata.normalize("NFKD", nome)
    nome = nome.encode("ascii", "ignore").decode("ascii")
    nome = re.sub(r"[^A-Za-z0-9._-]", "_", nome)
    return nome


def _mat_formatar_nome(texto):
    if not texto: return texto
    minusc = {"de", "da", "do", "das", "dos", "e"}
    return " ".join(p if p in minusc else p.capitalize()
                    for p in str(texto).strip().lower().split())


# ── Geração do Excel ──────────────────────────────────────────────────────────

COLUNAS_RENAME_MAT = {
    "matricula": "Matrícula", "comarca": "Comarca",
    "proprietario_atual": "Proprietário Atual", "descricao_imovel": "Descrição do Imóvel",
    "transmissoes_averbadas_registradas": "Transmissões",
    "onus_vigentes_registrados_averbados": "Ônus Vigentes",
    "observacoes": "Observações", "onus_cancelados": "Ônus Cancelados",
    "grau_confianca": "Grau de Confiança",
}
COLUNAS_ORDEM_MAT = [
    "Matrícula", "Comarca", "Proprietário Atual", "Descrição do Imóvel",
    "Transmissões", "Ônus Vigentes", "Observações", "Ônus Cancelados",
    "Fração Ideal", "Valor da Avaliação Definitiva (VM)",
    "Valor da Avaliação Definitiva (VP)", "Valor Total do Ônus",
    "Saldo Avaliação - Ônus", "Grau de Confiança",
]
LARGURAS_MAT = {1:12, 2:22, 3:30, 4:45, 5:55, 6:55, 7:50, 8:50, 9:14, 10:22, 11:22, 12:22, 13:22, 14:18}
COL_VM, COL_ONUS, COL_SALDO = 10, 12, 13

FILL_AMARELO  = PatternFill("solid", fgColor="FFFACD")  # lemon chiffon
FILL_VERMELHO = PatternFill("solid", fgColor="FFD0D0")  # rosa claro
FILL_PAR      = PatternFill("solid", fgColor="F2F7FB")
FILL_IMPAR    = PatternFill("solid", fgColor="FFFFFF")


def _mat_borda(estilo="thin", cor="BFBFBF"):
    s = Side(style=estilo, color=cor)
    return Border(left=s, right=s, top=s, bottom=s)


def _mat_estimar_linhas(texto, larg):
    if not texto: return 1
    chars = max(int(larg * 1.3), 1)
    total = 0
    for bloco in str(texto).split("\n"):
        total += 1 if bloco.strip() == "" else math.ceil(max(len(bloco), 1) / chars)
    return max(total, 1)


def _mat_salvar_excel(df, caminho, alert_map: dict):
    """
    alert_map: {df_index (0-based): set de alertas ('amarelo', 'vermelho')}
    """
    df.to_excel(caminho, index=False)
    wb = load_workbook(caminho)
    ws = wb.active
    ws.title = "Matriculas"
    n = len(df)
    ultima = 1 + n
    total  = ultima + 1
    lvm    = get_column_letter(COL_VM)
    lonus  = get_column_letter(COL_ONUS)

    for row in range(2, ultima + 1):
        ws.cell(row, COL_SALDO).value = f"={lvm}{row}-{lonus}{row}"
    ws.cell(total, 1).value = "TOTAL"
    for col in [COL_VM, COL_ONUS, COL_SALDO]:
        letra = get_column_letter(col)
        ws.cell(total, col).value = f"=SUM({letra}2:{letra}{ultima})"

    for col_idx, larg in LARGURAS_MAT.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = larg

    header_fill = PatternFill("solid", fgColor="1F4E79")
    ws.row_dimensions[1].height = 36
    for cell in ws[1]:
        cell.fill      = header_fill
        cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        cell.border    = _mat_borda("medium", "FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    colunas_center = {1, 2, 9, 10, 11, 12, 13, 14}
    for row_idx in range(2, ultima + 1):
        df_idx = row_idx - 2        # índice 0-based no DataFrame
        alertas = alert_map.get(df_idx, set())

        if "vermelho" in alertas:
            fill = FILL_VERMELHO
        elif "amarelo" in alertas:
            fill = FILL_AMARELO
        else:
            fill = FILL_PAR if row_idx % 2 == 0 else FILL_IMPAR

        max_linhas = 1
        for col_idx in range(1, ws.max_column + 1):
            cell  = ws.cell(row_idx, col_idx)
            horiz = "center" if col_idx in colunas_center else "left"
            cell.font      = Font(name="Calibri", size=10, color="1A1A1A")
            cell.fill      = fill
            cell.border    = _mat_borda("thin", "D0D0D0")
            cell.alignment = Alignment(horizontal=horiz, vertical="top", wrap_text=True)
            nl = _mat_estimar_linhas(cell.value, LARGURAS_MAT.get(col_idx, 20))
            max_linhas = max(max_linhas, nl)
        ws.row_dimensions[row_idx].height = min(max(16 * max_linhas + 6, 24), 409)

    total_fill = PatternFill("solid", fgColor="D6E4F0")
    ws.row_dimensions[total].height = 24
    for cell in ws[total]:
        cell.fill      = total_fill
        cell.font      = Font(bold=True, name="Calibri", size=10, color="1F4E79")
        cell.border    = _mat_borda("medium", "9DC3E6")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

    fmt_brl = r'R$\ #,##0.00'
    for col in [COL_VM, COL_VM + 1, COL_ONUS, COL_SALDO]:
        for row in range(2, total + 1):
            ws.cell(row, col).number_format = fmt_brl

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ultima}"

    # Legenda de cores (linha após o total)
    leg_row = total + 2
    ws.cell(leg_row, 1).value = "Legenda:"
    ws.cell(leg_row, 1).font = Font(bold=True, name="Calibri", size=9)
    ws.cell(leg_row, 2).fill = FILL_AMARELO
    ws.cell(leg_row, 2).value = "Transmissão envolvendo devedor ou grupo econômico"
    ws.cell(leg_row, 2).font = Font(name="Calibri", size=9)
    ws.cell(leg_row, 3).fill = FILL_VERMELHO
    ws.cell(leg_row, 3).value = "Transmissão após ajuizamento da execução"
    ws.cell(leg_row, 3).font = Font(name="Calibri", size=9)

    wb.save(caminho)


# ── Função pública ────────────────────────────────────────────────────────────

def mat_gerar_excel(arquivos,
                    data_ajuizamento: str = "",
                    devedores_txt: str = "",
                    relacionados_txt: str = ""):
    if not arquivos:
        yield "Envie ao menos um PDF.", "", None
        return

    # Parse dos parâmetros opcionais
    devedores_docs  = _mat_parse_parties(devedores_txt)
    relacionados_docs = _mat_parse_parties(relacionados_txt)
    data_ajuiz: date | None = None
    if data_ajuizamento and data_ajuizamento.strip():
        try:
            p = data_ajuizamento.strip().split("/")
            data_ajuiz = date(int(p[2]), int(p[1]), int(p[0]))
        except Exception:
            pass

    usar_alertas = bool(devedores_docs or relacionados_docs or data_ajuiz)

    try:
        client1, client2 = _mat_get_clients()
    except Exception as e:
        yield f"Erro de configuração: {e}", "", None
        return

    # Limpa arquivos temporários da execução anterior
    for f in os.listdir("tmp_pdfs"):
        try: os.remove(os.path.join("tmp_pdfs", f))
        except Exception: pass

    pdfs = []
    for arq in arquivos:
        origem = getattr(arq, "name", arq)
        nome   = _mat_limpar_nome(os.path.basename(origem))
        destino = os.path.join("tmp_pdfs", nome)
        with open(origem, "rb") as o, open(destino, "wb") as d:
            d.write(o.read())
        pdfs.append(destino)

    n = len(pdfs)
    log = [f"Analisando {n} matricula(s) em paralelo (4 workers)..."]
    if usar_alertas:
        partes = []
        if data_ajuiz:
            partes.append(f"ajuizamento: {data_ajuizamento.strip()}")
        if devedores_docs:
            partes.append(f"{len(devedores_docs)} devedor(es)")
        if relacionados_docs:
            partes.append(f"{len(relacionados_docs)} pessoa(s) do grupo")
        log.append(f"   Alertas ativos: {', '.join(partes)}")
    yield "\n".join(log), "", None

    resultados_dict: dict = {}
    erros_dict: dict     = {}

    # Prepara argumentos alternando entre os dois clientes
    args_list = [(i, pdf, client1 if i % 2 == 0 else client2) for i, pdf in enumerate(pdfs)]

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
        futures = {ex.submit(_mat_worker, args): args[0] for args in args_list}
        concluidos = 0
        for future in concurrent.futures.as_completed(futures):
            idx, linha, erro = future.result()
            concluidos += 1
            nome_pdf = Path(pdfs[idx]).name
            if erro:
                erros_dict[idx] = erro
                log.append(f"   [{concluidos}/{n}] {nome_pdf} — Erro: {erro}")
            else:
                resultados_dict[idx] = linha
                log.append(f"   [{concluidos}/{n}] {nome_pdf} — OK")
            yield "\n".join(log), "", None

    # Ordena por índice original
    indices_ok = sorted(resultados_dict.keys())
    hoje = date.today()
    alert_map: dict = {}

    resultados_finais = []
    for pos, idx in enumerate(indices_ok):
        r = resultados_dict[idx]
        onus_list = r.pop("onus_financeiros", None) or []
        onus_vigentes_txt = r.get("onus_vigentes_registrados_averbados", "") or ""
        r["valor_total_onus_calculado"] = _mat_calcular_valor_onus(
            onus_list, onus_vigentes_texto=onus_vigentes_txt, data_referencia=hoje
        )

        if usar_alertas:
            alertas = _mat_detectar_alertas(r, devedores_docs, relacionados_docs, data_ajuiz)
            if alertas:
                alert_map[pos] = alertas

        r.pop("transmissoes_estruturadas", None)  # não vai para o Excel
        resultados_finais.append(r)

    if not resultados_finais:
        log.append("\nNenhuma matrícula processada com sucesso.")
        yield "\n".join(log), "Nenhum resultado.", None
        return

    df = pd.DataFrame(resultados_finais)
    if "proprietario_atual" in df.columns:
        df["proprietario_atual"] = df["proprietario_atual"].apply(_mat_formatar_nome)
    df = df.rename(columns=COLUNAS_RENAME_MAT)
    df["Fração Ideal"] = ""
    df["Valor da Avaliação Definitiva (VM)"] = ""
    df["Valor da Avaliação Definitiva (VP)"] = ""
    df["Valor Total do Ônus"] = df.pop("valor_total_onus_calculado") if "valor_total_onus_calculado" in df.columns else ""
    df["Saldo Avaliação - Ônus"] = ""
    df = df[[c for c in COLUNAS_ORDEM_MAT if c in df.columns]]

    caminho = os.path.join("resultados", "resultado_matriculas.xlsx")
    _mat_salvar_excel(df, caminho, alert_map)

    linhas_alerta = []
    if alert_map:
        n_am = sum(1 for a in alert_map.values() if "amarelo" in a and "vermelho" not in a)
        n_vm = sum(1 for a in alert_map.values() if "vermelho" in a)
        if n_am: linhas_alerta.append(f"{n_am} matrícula(s) amarela(s)")
        if n_vm: linhas_alerta.append(f"{n_vm} matrícula(s) vermelha(s)")

    resumo = f"{len(resultados_finais)}/{n} OK"
    if linhas_alerta:
        resumo += " | " + ", ".join(linhas_alerta)
    if erros_dict:
        resumo += f" | {len(erros_dict)} erro(s)"

    log.append(f"\nExcel gerado — {resumo}")
    yield "\n".join(log), "Excel pronto!", caminho


def mat_responder(pergunta: str, log_texto: str):
    try:
        k = os.getenv("GEMINI_API_KEY_1") or os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k)
        return _responder_pergunta_generica(
            pergunta, log_texto, client, "gemini-2.5-flash"
        )
    except Exception as e:
        return f"Erro: {e}"
