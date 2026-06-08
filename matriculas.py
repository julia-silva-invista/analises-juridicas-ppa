# -*- coding: utf-8 -*-
import os
import re
import math
import unicodedata
import time
import tempfile
from datetime import date
from pathlib import Path
from typing import Optional, List

from dateutil.relativedelta import relativedelta

import pandas as pd
from pydantic import BaseModel
from google import genai

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from utils import _retry, _responder_pergunta_generica
_mat_api_key = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
_mat_client = genai.Client(api_key=_mat_api_key) if _mat_api_key else None

class OnusFinanceiro(BaseModel):
    codigo: Optional[str] = None           # ex: "R.5", "AV.9"
    valor_principal: Optional[float] = None  # em reais
    data_celebracao: Optional[str] = None  # DD/MM/AAAA
    parcela_mensal: Optional[float] = None  # em reais, se disponível
    vencimento_final: Optional[str] = None  # DD/MM/AAAA, se disponível


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


PROMPT_MATRICULA = (
    "Voce e um assistente especializado em analise de matriculas de imoveis brasileiros.\n\n"
    "Analise integralmente o PDF enviado e extraia as informacoes em formato estruturado.\n\n"
    "Preencha os seguintes campos:\n"
    "- matricula\n- comarca\n- proprietario_atual\n- descricao_imovel\n"
    "- transmissoes_averbadas_registradas\n- onus_vigentes_registrados_averbados\n"
    "- observacoes\n- onus_cancelados\n- grau_confianca\n\n"
    "Regras obrigatorias:\n\n"
    "1. Proprietario atual: ultimo registro translativo valido. APENAS nomes, sem CPF/CNPJ.\n"
    "   Se a matrícula estiver encerrada, indique o proprietario e registre em observacoes.\n\n"
    "   Se a matrícula pertencer a mais de uma pessoa, indique o percentual correspondente a propriedade de cada um dos titulares atuais. por exemplo: Fulano da Silva (40%) Beltrano Camargo (60%)\n\n"
    "2. comarca: indicar a comerca e qual é o cartório. Exemplo '1º CRI de Curitiba/PR'.\n\n"
    "3. descricao_imovel: resumo em ate 3 frases. Tipo, localizacao, area, identificadores.\n\n"
    "4. transmissoes_averbadas_registradas: todas, cronologicas, sinteticas, com CPF/CNPJ (nao precisa ficar repetindo se ja colocou uma vez). Diga sempre o preço do imóvel em caso de compra e venda ou integralização.\n"
    "   Inclua registros anteriores a abertura se disponiveis.\n"
    "   Exemplo: R.4 - Compra e Venda: Jose Marques da Cruz vendeu o imóvel para Carlos Guimaraes em 15/12/1986, pelo valor declarado de R$ 2.000.000,00\n"
    "   E sempre que trouxer a natureza das transmissões, coloque em negrito apenas o tipo de transmissão, por exemplo Compra e Venda ou Partilha.\n\n"
    "5. onus_vigentes_registrados_averbados: apenas gravames vigentes.\n"
    "   Se nao houver, preencha exatamente: Sem onus vigentes identificados.\n\n"
    "6. Formato de cada onus:\n"
    "   [Codigo]-[Matrícula]: [Tipo], [n. cedula se disponivel], [dd/mm/aaaa], [partes sem CPF/CNPJ],\n"
    "   vencimento [data], Valor principal: R$ [valor se disponivel].\n"
    "   NAO inclua vara/comarca/processo nos onus (vai em observacoes).\n\n"
    "7. Nao inclua atos cancelados em onus_vigentes. Se cancelado, mencione em onus_cancelados.\n\n"
    "8. PDF escaneado: extraia o maximo possivel do conteudo visual.\n\n"
    "9. Nunca invente dados. Nomes: primeira letra maiuscula, exceto artigos (de, da, do, das, dos, e).\n"
    "   Erros de grafia: copie como esta e registre em observacoes.\n\n"
    "10. observacoes: inclua SEMPRE que houver:\n"
    "   - Vara, comarca e n. processo de penhoras\n"
    "   - Averbacoes relevantes (divorcio, matrimonio, georreferenciamento, reserva legal, quitacao)\n"
    "   - Irregularidades, CPFs divergentes para a mesma pessoa, erros de grafia\n"
    "   - Matrícula encerrada, rematricula, baixa legibilidade\n"
    "   Se nao houver: Sem observacoes relevantes. Nunca deixe vazio.\n\n"
    "11. onus_cancelados: inclua APENAS onus cancelados, sempre indicando, se possivel, qual ato os cancelou.\n"
    "    Exemplo: AV.9: Penhora cancelada pela AV.22.\n"
    "    Se nao houver: Sem onus cancelados identificados.\n\n"
    "12. grau_confianca: Alto, Medio ou Baixo.\n"
    "    Medio/Baixo: justificativa entre parenteses, indique quais registros e avs nao podem ser lidos, se for o caso. Ex: Medio (Baixa legibilidade da AV-3).\n"
    "    Alto: preencha apenas Alto.\n"
    "    A marca d'agua 'nao possui valor de certidao' por si so nao deve ser considerada indicio de medio ou baixa confianca.\n\n"
    "    ATENCAO: matricula encerrada ou rematriculada = obrigatoriamente Medio ou Baixo.\n\n"
    "13. Responda apenas com JSON valido aderente ao schema.\n\n"
    "14. Redacao padronizada e objetiva. Evite textos longos.\n\n"
    "15. Pule UMA LINHA entre itens autonomos diferentes. Exemplo:\n\n"
    "    R.1 - Compra e Venda: Jose Marques da Cruz vendeu o imóvel para Carlos Guimaraes em 15/12/1986, pelo valor declarado de R$ 2.000.000,00.\n\n"
    "    R.2 - Doação: Carlos Guimaraes doou o imóvel para Murilo Menezes em 03/06/1988.\n\n"
    "    R.3 - Integralizacao: Murilo Menezes integralizou o imóvel na sociedade Fluxo Invest Ltda, pelo valor declarado de R$2.500.000,00, em 01/11/2000.\n\n"
    "16. NUNCA pule linha dentro de uma mesma frase. Intervalos como AV.1 a AV.8 ou "
    "referencias como conforme R.5 sao parte da frase e NAO devem ter quebra de linha.\n\n"
    "17. Apos a analise, verifique todos os AVs/Rs para ver se algum ficou faltando. "
    "(exemplo, se vai ate o R-40 ou AV.40, leia tudo e veja se todas AVs/Rs foram analisados e mapeados. "
    "se for, por exemplo, ate o r-4, devera conferir se o R.1, R.2, R.3 e R.4 foram mapeados). "
    "Caso falte alguma averbaçao, indique em 'observacoes' qual e o registro ou averbaçao cujo conteudo nao pode ser identificado.\n\n"
    "18. Para cada onus vigente de natureza financeira (penhora em execucao, hipoteca, alienacao fiduciaria, "
    "CCB, confissao de divida ou qualquer titulo de credito), preencha o campo 'onus_financeiros' com:\n"
    "   - codigo: codigo do registro/averbaçao (ex: R.5, AV.9)\n"
    "   - valor_principal: valor principal em reais (numero puro, sem R$ ou pontos)\n"
    "   - data_celebracao: data do ato em formato DD/MM/AAAA\n"
    "   - parcela_mensal: valor da parcela mensal em reais, SE explicitamente indicada no documento "
    "(apenas se constar 'parcelas mensais de R$ X' ou equivalente). Deixe null se nao houver.\n"
    "   - vencimento_final: data de vencimento final em DD/MM/AAAA, SE disponivel. Deixe null se nao houver.\n"
    "   Nao inclua onus sem valor principal identificavel. "
    "Nao invente dados — se nao tiver certeza, deixe o campo null.\n\n"
    "19. VARREDURA OBRIGATORIA AV/R POR AV/R:\n"
    "    a) Identifique a numeracao maxima de R e de AV no documento (ex: R.7, AV.22).\n"
    "    b) Percorra individualmente de R.1 ate o ultimo R, e de AV.1 ate o ultimo AV.\n"
    "    c) Para CADA item, classifique em EXATAMENTE uma destas categorias:\n"
    "       - ONUS VIGENTE: penhora ativa, hipoteca, alienacao fiduciaria, CCB, arresto vigente, "
    "confissao de divida, qualquer titulo de credito ainda nao cancelado "
    "-> vai em onus_vigentes_registrados_averbados\n"
    "       - ONUS CANCELADO: qualquer onus cuja AV/R posterior tenha cancelado/baixado/extinto "
    "-> vai em onus_cancelados (sempre indicando qual ato cancelou, ex: 'AV.9: Penhora cancelada pela AV.22')\n"
    "       - OBSERVACAO/TRANSMISSAO: compra e venda, doacao, partilha, integralizacao, divorcio, "
    "retificacao, georreferenciamento, encerramento, especificacao, anuencia, etc. "
    "-> vai em transmissoes_averbadas_registradas (se transmissao) ou observacoes (demais casos)\n"
    "    d) NUNCA pule nenhum R/AV. Se algum estiver ilegivel ou nao puder ser identificado, "
    "cite-o explicitamente em observacoes (ex: 'AV.X nao pode ser lida'). "
    "A omissao de qualquer R/AV do documento e ERRO GRAVE.\n"
    "    e) Antes de finalizar a resposta, confira se TODOS os R/AV de R.1/AV.1 ate o ultimo "
    "aparecem em pelo menos uma das tres categorias acima ou em observacoes.\n\n"
)

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


def _mat_calcular_valor_onus(onus_list: list, data_referencia: date = None) -> str:
    if not onus_list:
        return ""
    if data_referencia is None:
        data_referencia = date.today()

    linhas_individuais = []
    total = 0.0

    for onus in onus_list:
        codigo = onus.get("codigo") if isinstance(onus, dict) else getattr(onus, "codigo", None)
        valor = onus.get("valor_principal") if isinstance(onus, dict) else getattr(onus, "valor_principal", None)
        data_str = onus.get("data_celebracao") if isinstance(onus, dict) else getattr(onus, "data_celebracao", None)
        parcela = onus.get("parcela_mensal") if isinstance(onus, dict) else getattr(onus, "parcela_mensal", None)

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

        prefixo = f"∘ {codigo}: " if codigo else "∘ "
        linha = prefixo + _fmt_brl(valor_atualizado)

        if parcela and parcela > 0:
            saldo = valor_atualizado - parcela * meses
            if saldo > 0:
                linha += f"\n   (em caso de adimplência, saldo de {_fmt_brl(saldo)})"

        linhas_individuais.append(linha)

    if not linhas_individuais:
        return ""

    cabecalho = f"Total: {_fmt_brl(total)}"
    return cabecalho + "\n" + "\n".join(linhas_individuais)


def _mat_pos_processar(resultado):
    campos = ["proprietario_atual", "descricao_imovel", "transmissoes_averbadas_registradas",
              "onus_vigentes_registrados_averbados", "observacoes", "onus_cancelados"]
    for campo in campos:
        v = resultado.get(campo)
        if v:
            v = _mat_remover_cpf(v)
            v = _mat_corrigir_caps(v)
            v = _mat_normalizar_siglas(v)
            resultado[campo] = v
    for campo in ["transmissoes_averbadas_registradas", "onus_vigentes_registrados_averbados",
                  "observacoes", "onus_cancelados"]:
        v = resultado.get(campo)
        if v:
            v = _mat_normalizar_quebras(v)
            v = _mat_normalizar_siglas(v)
            resultado[campo] = v
    return resultado

def _mat_analisar_pdf(caminho_pdf: str):
    if not _mat_client:
        raise RuntimeError("GOOGLE_API_KEY nao configurada.")

    try:
        caminho_pdf.encode("ascii")
        upload_path = caminho_pdf
    except UnicodeEncodeError:
        import shutil as _shutil
        tmp_ascii = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_ascii.close()
        _shutil.copy2(caminho_pdf, tmp_ascii.name)
        upload_path = tmp_ascii.name

    arquivo = _mat_client.files.upload(file=upload_path)

    def _call():
        response = _mat_client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[PROMPT_MATRICULA, arquivo],
            config={
                "response_mime_type": "application/json",
                "response_json_schema": MatriculaExtraida.model_json_schema(),
            },
        )
        return MatriculaExtraida.model_validate_json(response.text).model_dump()

    resultado = _retry(_call, tentativas=3, espera_base=10)
    return _mat_pos_processar(resultado)

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

def _mat_salvar_excel(df, caminho):
    df.to_excel(caminho, index=False)
    wb = load_workbook(caminho)
    ws = wb.active
    ws.title = "Matriculas"
    n = len(df)
    ultima = 1 + n
    total = ultima + 1
    lvm = get_column_letter(COL_VM)
    lonus = get_column_letter(COL_ONUS)
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
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        cell.border = _mat_borda("medium", "FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    colunas_center = {1, 2, 9, 10, 11, 12, 13, 14}
    for row_idx in range(2, ultima + 1):
        par = row_idx % 2 == 0
        fill = PatternFill("solid", fgColor="F2F7FB" if par else "FFFFFF")
        max_linhas = 1
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row_idx, col_idx)
            horiz = "center" if col_idx in colunas_center else "left"
            cell.font = Font(name="Calibri", size=10, color="1A1A1A")
            cell.fill = fill
            cell.border = _mat_borda("thin", "D0D0D0")
            cell.alignment = Alignment(horizontal=horiz, vertical="top", wrap_text=True)
            nl = _mat_estimar_linhas(cell.value, LARGURAS_MAT.get(col_idx, 20))
            max_linhas = max(max_linhas, nl)
        ws.row_dimensions[row_idx].height = min(max(16 * max_linhas + 6, 24), 409)
    total_fill = PatternFill("solid", fgColor="D6E4F0")
    ws.row_dimensions[total].height = 24
    for cell in ws[total]:
        cell.fill = total_fill
        cell.font = Font(bold=True, name="Calibri", size=10, color="1F4E79")
        cell.border = _mat_borda("medium", "9DC3E6")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
    fmt_brl = r'R$\ #,##0.00'
    for col in [COL_VM, COL_VM + 1, COL_ONUS, COL_SALDO]:
        for row in range(2, total + 1):
            ws.cell(row, col).number_format = fmt_brl
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ultima}"
    wb.save(caminho)

def mat_gerar_excel(arquivos):
    if not arquivos:
        yield "Envie ao menos um PDF.", "", None
        return

    # Limpa arquivos temporários da execução anterior
    for f in os.listdir("tmp_pdfs"):
        try:
            os.remove(os.path.join("tmp_pdfs", f))
        except Exception:
            pass

    pdfs = []
    for arq in arquivos:
        origem = getattr(arq, "name", arq)
        nome = _mat_limpar_nome(os.path.basename(origem))
        destino = os.path.join("tmp_pdfs", nome)
        with open(origem, "rb") as o, open(destino, "wb") as d:
            d.write(o.read())
        pdfs.append(destino)

    log = [f"Analisando {len(pdfs)} matricula(s)..."]
    yield "\n".join(log), "", None

    resultados = []
    for i, pdf in enumerate(pdfs, 1):
        log.append(f"   [{i}/{len(pdfs)}] {Path(pdf).name}...")
        yield "\n".join(log), "", None
        try:
            linha = _mat_analisar_pdf(pdf)
            resultados.append(linha)
            log[-1] = f"   [{i}/{len(pdfs)}] {Path(pdf).name} — OK"
        except Exception as e:
            log[-1] = f"   [{i}/{len(pdfs)}] Erro: {e}"
        yield "\n".join(log), "", None

    hoje = date.today()
    for r in resultados:
        onus_list = r.pop("onus_financeiros", None) or []
        r["valor_total_onus_calculado"] = _mat_calcular_valor_onus(onus_list, hoje)

    df = pd.DataFrame(resultados)
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
    _mat_salvar_excel(df, caminho)
    log.append("\nExcel gerado com sucesso!")
    yield "\n".join(log), "Excel pronto!", caminho


def mat_responder(pergunta: str, log_texto: str):
    try:
        k = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY_1") or os.getenv("GEMINI_API_KEY")
        client = genai.Client(api_key=k)
        return _responder_pergunta_generica(
            pergunta, log_texto, client, "gemini-2.5-flash"
        )
    except Exception as e:
        return f"Erro: {e}"


