# -*- coding: utf-8 -*-
"""Regras compartilhadas de fidelidade e rastreabilidade jurídica.

Estas regras são acrescentadas aos prompts existentes. Elas não substituem o
conteúdo jurídico ou o formato de cada relatório; apenas tornam explícita a
origem de cada fato e impedem que marcadores internos de processamento sejam
tratados como referências dos autos.
"""

import re


REGRA_FIDELIDADE_PROCESSUAL = """\
======================================================================
PROTOCOLO DE FIDELIDADE E RASTREABILIDADE — REGRA ABSOLUTA
======================================================================
1. Separe sempre três conceitos:
   a) IDENTIFICADOR REAL DO TRIBUNAL: ID, Mov., Evento ou numeração de folhas dos
      autos exibida pelo próprio sistema/peça;
   b) PÁGINA ABSOLUTA DO PDF: posição física, iniciada em 1, dentro do arquivo
      original enviado pela usuária;
   c) MARCADOR INTERNO: PARTE, chunk, lote, página local do trecho, texto bruto ou
      qualquer separador criado apenas pelo robô.
   Somente (a) e (b) podem aparecer como referência no relatório. O item (c)
   JAMAIS pode aparecer na resposta final.

2. Toda afirmação factual relevante deve ser rastreável ao identificador REAL
   do tribunal e à página ABSOLUTA do PDF em que o dado foi efetivamente lido.
   Copie ID/Mov./Evento/fls. exatamente como aparece. Nunca estime um identificador
   a partir da página anterior ou seguinte e nunca transforme o número de página
   do PDF em ID, Mov., Evento ou folha interna dos autos.

3. Formato com UM único PDF analisado:
   (ID 188753786 | fl. 135)
   (Mov. 42.1 | fl. 27)
   (Evento 54 | fls. 65/66)
   Se o tribunal numera os próprios autos por folhas e essa numeração diverge da
   página física: (fls. 77/94 TJSP | fls. 65/70)
   Com um único PDF, é PROIBIDO escrever "do pdf", o nome do arquivo, "parte X",
   "página do texto bruto" ou qualquer equivalente.

4. Formato quando MAIS DE UM PDF foi analisado:
   (ID 188753786 | fl. 135 do pdf P3.PDF)
   (Evento 54 | fls. 65/66 do pdf recurso.pdf)
   Escreva sempre "pdf" em minúsculas. Use o nome real do arquivo ou, quando
   inequivocamente identificado, o número do processo correspondente. Não use o
   número do chunk/parte como nome do pdf.

5. A página citada é sempre a página absoluta do PDF original, nunca a posição
   dentro do chunk. Se a fonte disser [PDF_PAGE: 425], cite fl. 425, ainda que o
   trecho enviado ao modelo comece localmente na página 1.

6. Se o fato estiver legível, mas o identificador do tribunal não estiver, não
   invente. Registre na extração "IDENTIFICADOR_TRIBUNAL: não localizado" e
   preserve a página absoluta. Na redação final, explicite a limitação como
   "identificador processual não localizado"; nunca apresente apenas uma página
   interna como se fosse referência completa do tribunal.

7. Antes de finalizar, confira cada data, valor, índice, cláusula, nome, evento e
   referência contra a mesma evidência de origem. Uma referência existente, mas
   ligada a outro fato ou a outra página, também é erro grave.
"""


REGRA_EXTRACAO_POR_PAGINA = """\
======================================================================
PROTOCOLO DE EXTRAÇÃO PÁGINA A PÁGINA
======================================================================
- Organize a extração em blocos de evidência. Para cada página que contenha
  informação útil, abra o bloco com:
  [FONTE_PDF: nome real | PDF_PAGE: número absoluto | IDENTIFICADOR_TRIBUNAL:
  ID/Mov./Evento/fls. exato ou "não localizado"]
- Dentro do bloco, transcreva datas, valores, índices, cláusulas, partes e atos
  exatamente como constam. Não antecipe a síntese jurídica e não corrija datas
  porque a ordem física do processo parece estranha.
- Um documento pode continuar por várias páginas. Mantenha um bloco por página e
  repita o identificador quando ele estiver visível. Não atribua automaticamente
  a uma página o ID/Mov./Evento lido em outra.
- PARTE, chunk, lote e página local servem apenas para transporte. Nunca os use
  como fonte, andamento, evento ou referência jurídica.
"""


REGRA_INDICES_E_ADITAMENTOS = """\
======================================================================
CONTRATOS, ÍNDICES E ADITAMENTOS — REGRA ABSOLUTA
======================================================================
- Leia o contrato/lastro e todos os aditamentos visualmente, inclusive quando
  digitalizados. Não deduza índice contratual a partir da petição inicial, de uma
  planilha ou de decisão posterior.
- Em "a) Correção monetária", escreva o índice exato previsto no contrato e a
  cláusula que o disciplina. Se o contrato não contiver índice de correção
  monetária, escreva exatamente "Não há". Não use "Não especificado" quando a
  leitura integral permitir concluir que não há previsão.
- Para correção monetária, juros remuneratórios, juros moratórios, multa e
  capitalização, informe: índice/taxa, periodicidade, base de cálculo quando
  houver, número/título da cláusula e referência processual completa com página
  absoluta do PDF.
- Diferencie sempre: (i) índices do contrato original; (ii) alterações de cada
  aditamento; (iii) índices efetivamente usados na planilha inicial; (iv) índices
  da memória de cálculo mais recente; e (v) critérios fixados por decisão judicial.
  Não misture as fontes.
- Mapeie TODOS os aditamentos principais. Se houver dúvida razoável sobre a
  relevância, inclua o aditamento. Para cada um, informe data, instrumento afetado,
  cláusulas/condições alteradas, partes/assinaturas e referência completa.
"""


REGRA_CRONOLOGIA_PROCESSUAL = """\
======================================================================
PRINCIPAIS ANDAMENTOS — COMPLETUDE E ORDEM CRONOLÓGICA
======================================================================
- Inclua todos os andamentos principais. Se houver dúvida razoável sobre a
  relevância de um ato, inclua-o; é preferível manter um ato potencialmente
  relevante a omitir um marco que altere a compreensão do processo.
- Primeiro extraia a DATA DO ATO de cada evidência; só depois ordene a lista.
  A ordem física das páginas, dos IDs, dos movimentos, dos eventos, dos arquivos
  ou dos chunks não substitui a cronologia.
- Ordene em ordem cronológica crescente pela data real do ato. Se 01/2002 aparecer
  fisicamente antes de 10/2000, 10/2000 deve vir primeiro no relatório. Essa
  reorganização jamais autoriza trocar as referências: cada ato conserva o ID,
  Mov., Evento/fls. e a página absoluta do PDF em que foi encontrado.
- Não corrija silenciosamente uma data conflitante. Se duas fontes trouxerem datas
  incompatíveis para o mesmo ato, registre a divergência e cite ambas.
- Antes de finalizar, faça uma varredura de regressões temporais e confirme que
  nenhum período ou PDF ficou sem cobertura.
"""


REGRA_AUDITORIA_FINAL = """\
======================================================================
AUDITORIA FINAL OBRIGATÓRIA
======================================================================
Antes de entregar:
1. compare cada data, valor, índice, cláusula e referência com a extração;
2. confirme que não há "PARTE X", "chunk", "texto bruto", página local ou página
   de processamento na resposta;
3. confirme que toda página citada é absoluta no PDF original;
4. confirme que o sufixo "do pdf ..." existe somente quando foram analisados dois
   ou mais PDFs;
5. confira a ordem cronológica dos principais andamentos;
6. confira todos os aditamentos e a cláusula de cada índice contratual;
7. se a evidência não sustentar uma afirmação, remova a afirmação ou sinalize a
   incerteza — jamais complete por plausibilidade jurídica.
"""


REGRAS_CONSOLIDACAO_PROCESSUAL = (
    REGRA_FIDELIDADE_PROCESSUAL
    + "\n"
    + REGRA_INDICES_E_ADITAMENTOS
    + "\n"
    + REGRA_CRONOLOGIA_PROCESSUAL
    + "\n"
    + REGRA_AUDITORIA_FINAL
)


REGRA_REFERENCIA_MATRICULA = """\
======================================================================
RASTREABILIDADE DA MATRÍCULA
======================================================================
- A matrícula não possui necessariamente ID/Mov./Evento de tribunal. Use como
  identificador primário o código registral real (R., AV., matrícula anterior,
  certidão ou outro ato que apareça no documento) e acrescente a página absoluta
  do PDF original: (R.5 | fl. 8), (AV.12 | fls. 14/15).
- Nunca cite PARTE, chunk, texto bruto ou página local de processamento.
- Se mais de um PDF compuser a mesma análise, use "fl. X do pdf nome.pdf"; com um
  único PDF, não escreva "do pdf".
- Cada transmissão, ônus, cancelamento, fração e encerramento deve conservar o
  código registral e a página em que foi efetivamente lido. Não associe a um ato
  a página ou o código de outro ato.
- Ordene os atos pela data registral quando ela existir, sem perder a referência
  correta. Se a disposição física não for cronológica, reorganize a apresentação,
  não a evidência.
"""


def bloco_instrucao_adicional(texto: str, permite_secao_nova: bool = True) -> str:
    """Instrução do usuário como seção nomeada, com precedência sobre o template.

    Antes o texto entrava como um parágrafo solto no fim do prompt, competindo com um
    template que enumera as seções e manda segui-lo à risca — pedir "abra uma seção
    nova" era simplesmente ignorado. Aqui a instrução é rotulada e a autorização é
    explícita, sem afrouxar a regra que importa: continua proibido inventar fato.
    """
    pedido = str(texto or "").strip()
    if not pedido:
        return ""

    regras = [
        "- Atenda ao pedido acima integralmente, além de tudo o que o template já exige.",
        "- O pedido NÃO autoriza inventar, estimar ou presumir fato algum. Se o material "
        "não trouxer o que foi pedido, diga expressamente que não consta.",
        "- Toda informação trazida por causa do pedido leva referência (fls./Mov./Evento/ID) "
        "como qualquer outra.",
    ]
    if permite_secao_nova:
        regras.append(
            "- Se o pedido não couber em nenhum item previsto, ACRESCENTE um item novo ao "
            "fim da subseção do processo, com título próprio, em vez de espremer a "
            "informação num item que não é dela ou de omiti-la."
        )

    return (
        "═══ INSTRUÇÃO ADICIONAL DA ANALISTA (PRIORITÁRIA) ═══\n"
        f"{pedido}\n\n"
        + "\n".join(regras)
        + "\n\n"
    )


def contexto_fonte_pdf(nome_arquivo: str, pagina_inicial: int, pagina_final: int,
                       paginas_digitalizadas=None) -> str:
    """Contexto determinístico que vincula páginas locais às páginas originais."""
    digitalizadas = sorted({int(p) for p in (paginas_digitalizadas or [])})
    scan = ", ".join(str(p) for p in digitalizadas) if digitalizadas else "nenhuma detectada"
    return (
        "CONTEXTO TÉCNICO DA FONTE — NÃO CITE ESTE CABEÇALHO NO RELATÓRIO:\n"
        f"- arquivo original: {nome_arquivo}\n"
        f"- páginas absolutas cobertas: {pagina_inicial}-{pagina_final}\n"
        f"- a página local 1 deste trecho corresponde à página absoluta {pagina_inicial} do PDF original\n"
        f"- páginas digitalizadas detectadas neste trecho: {scan}\n"
        "Use esses números somente para preencher PDF_PAGE nos blocos de evidência. "
        "Nunca use PARTE/chunk/página local como referência jurídica.\n"
    )


_REF_INTERNA_RE = re.compile(
    r"\((?=[^()]{0,180}(?:texto\s+bruto|chunk\s*\d*|parte\s+\d+(?:\s*/\s*\d+)?|p[aá]gina\s+local))"
    r"[^()]*\)",
    re.IGNORECASE,
)
_LINHA_INTERNA_RE = re.compile(
    r"(?im)^\s*(?:=+\s*)?(?:PARTE\s+\d+(?:\s*/\s*\d+)?|CHUNK\s+\d+)[^\n]*\n?"
)
_REF_SOMENTE_PAGINA_RE = re.compile(
    r"\(\s*(fls?\.\s*\d+(?:\s*[/,-]\s*\d+)*)\s*\)", re.IGNORECASE
)
_REF_PAGINA_TECNICA_RE = re.compile(
    r"\(\s*(?:p\.|p[aá]g(?:ina)?\.?)\s*\d+(?:\s*[/,-]\s*\d+)*\s*\)",
    re.IGNORECASE,
)


def normalizar_referencias_relatorio(texto: str, multiplos_pdfs: bool) -> str:
    """Barreira final contra referências técnicas que vazem do processamento.

    Não tenta inventar a referência correta. Quando o modelo devolver apenas uma
    página ou um marcador interno, torna a limitação explícita para que o número
    técnico não seja apresentado como ID/Mov./Evento real.
    """
    texto = str(texto or "")
    texto = _LINHA_INTERNA_RE.sub("", texto)
    texto = _REF_INTERNA_RE.sub("(referência processual não localizada)", texto)
    texto = _REF_PAGINA_TECNICA_RE.sub("(referência processual não localizada)", texto)
    texto = _REF_SOMENTE_PAGINA_RE.sub(
        lambda m: f"(identificador processual não localizado | {m.group(1)})",
        texto,
    )
    texto = re.sub(r"\bPDF\b", "pdf", texto)
    if not multiplos_pdfs:
        # Em referência parentética de análise com arquivo único, o nome do pdf é
        # redundante e expressamente proibido. Mantém todos os demais elementos.
        texto = re.sub(
            r"(\([^()]*?)\s+do\s+pdf\s+[^()|;]+(?=\))",
            r"\1",
            texto,
            flags=re.IGNORECASE,
        )
    return re.sub(r"\n{3,}", "\n\n", texto).strip()


def problemas_referencia_interna(texto: str) -> list[str]:
    """Lista vazamentos que nunca podem chegar ao relatório final."""
    texto = str(texto or "")
    problemas = []
    if _REF_INTERNA_RE.search(texto) or _LINHA_INTERNA_RE.search(texto):
        problemas.append("marcador interno de parte/chunk/texto bruto")
    if re.search(r"\bp[aá]gina\s+(?:do\s+)?texto\s+bruto\b", texto, re.IGNORECASE):
        problemas.append("página do texto bruto")
    return problemas


def normalizar_referencias_objeto(valor, multiplos_pdfs: bool):
    """Aplica a barreira final recursivamente a JSONs de dossiê/checklist."""
    if isinstance(valor, dict):
        return {
            chave: normalizar_referencias_objeto(conteudo, multiplos_pdfs)
            for chave, conteudo in valor.items()
        }
    if isinstance(valor, list):
        return [normalizar_referencias_objeto(item, multiplos_pdfs) for item in valor]
    if isinstance(valor, str):
        return normalizar_referencias_relatorio(valor, multiplos_pdfs)
    return valor
