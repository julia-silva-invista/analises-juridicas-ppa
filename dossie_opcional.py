# -*- coding: utf-8 -*-
"""Capítulo "Opcional" do Dossiê Prévia — o que só existia no Dossiê PPA.

Os dois dossiês nascem da MESMA extração (`dossie_ppa._extrair_dados`): a Prévia
sempre teve os dados do PPA em mãos e simplesmente não os mostrava, porque o
template dela é de triagem. Este módulo devolve esse material à Prévia, num
capítulo à parte, seguindo três regras:

1. **Nada duplicado.** Campo que a Prévia já exibe nos quadros dela (número do
   processo, executados, SOP/SAT, lastro, garantia, status, constrições, imóveis,
   riscos) não se repete aqui.
2. **Nada em branco.** Quadro cuja informação a análise não preencheu não é
   desenhado. Passivo fiscal/trabalhista/cível, pendências e folha de revisão do
   PPA dependem de e-CAC, certidões e diligências externas ao processo — ficam de
   fora por isso, e não por esquecimento.
3. **Um bloco por crédito.** Com mais de uma execução no caso, a visão jurídica é
   replicada e numerada por crédito, no mesmo par credor + número de processo que
   nomeia a seção principal.
"""

from __future__ import annotations

from dossie_ppa import (
    _body,
    _compactar_celula,
    _filtrar_ativos,
    _grid_table,
    _kv_label_table,
    _kv_table,
    _note,
    _para,
    _somar_moeda,
    _sub_gray,
    _sub_orange,
    _spacer,
    _texto_analise,
    _LARANJA,
    _TXT,
    lastros_do_credito,
    rotulo_credito,
)

TITULO_CAPITULO = "6. OPCIONAL"

# Larguras do template oficial do PPA, para o quadro de ativos sair igual ao de lá.
_W_ATIVOS = [1.48, 1.78, 2.22, 2.26, 2.16, 1.47, 1.19, 1.19, 1.36, 1.40]
_COLS_ATIVOS = [
    "Mat.", "Comarca", "Proprietário Atual", "Descrição do Imóvel", "Ônus Vigentes",
    "Fração", "VM (R$)", "VF (R$)", "Ônus Total (R$)", "Saldo (R$)",
]


# ── Leitura defensiva do JSON da extração ─────────────────────────────────────

def _txt(valor) -> str:
    return str(valor or "").strip()


def _dicts(valor) -> list:
    return [item for item in (valor or []) if isinstance(item, dict)] if isinstance(valor, list) else []


def _dict(valor) -> dict:
    return valor if isinstance(valor, dict) else {}


def _pares_preenchidos(pares) -> list:
    """Descarta os rótulos sem valor — quadro só com rótulo não informa nada."""
    return [(rotulo, _txt(valor)) for rotulo, valor in pares if _txt(valor)]


# ── Blocos ────────────────────────────────────────────────────────────────────

def _quadro_kv(doc, titulo: str, pares) -> bool:
    linhas = _pares_preenchidos(pares)
    if not linhas:
        return False
    _sub_gray(doc, titulo)
    _kv_table(doc, ["CAMPO", "INFORMAÇÃO"], linhas)
    _spacer(doc, pts=2)
    return True


def _quadro_grade(doc, titulo: str, colunas, larguras, registros) -> bool:
    linhas = [linha for linha in registros if any(_txt(c) for c in linha)]
    if not linhas:
        return False
    _sub_gray(doc, titulo)
    _grid_table(doc, colunas, linhas, larguras)
    _spacer(doc, pts=2)
    return True


def _quadro_ativos(doc, titulo: str, ativos) -> bool:
    itens = _dicts(ativos)
    if not itens:
        return False
    linhas = [[
        item.get("matricula", ""),
        item.get("comarca", ""),
        item.get("proprietario_atual", ""),
        item.get("descricao") or item.get("tipo_ativo", ""),
        item.get("onus_vigentes", ""),
        item.get("fracao_atingivel", ""),
        item.get("vm", ""),
        item.get("vp", ""),
        item.get("onus_total", ""),
        item.get("saldo", ""),
    ] for item in itens]
    total = [
        "TOTAL", "", "", "", "", "",
        _somar_moeda(itens, "vm"),
        _somar_moeda(itens, "vp"),
        _somar_moeda(itens, "onus_total"),
        _somar_moeda(itens, "saldo"),
    ]
    _sub_gray(doc, titulo)
    tabela = _grid_table(doc, _COLS_ATIVOS, linhas, _W_ATIVOS, total_row=total)
    # Dez colunas em 16,9 cm: sem reduzir a fonte, a tabela quebra a linha em toda célula.
    for linha in tabela.rows[1:]:
        for indice, celula in enumerate(linha.cells):
            _compactar_celula(celula, 5.3 if indice >= 6 else 6.2, centralizar=indice >= 5)
    _spacer(doc, pts=2)
    return True


def _texto_livre(doc, titulo: str, valor) -> bool:
    texto = _texto_analise(valor)
    if not texto:
        return False
    _sub_gray(doc, titulo)
    for paragrafo in [p for p in texto.split("\n") if p.strip()]:
        _body(doc, paragrafo)
    _spacer(doc, pts=2)
    return True


# ── 6.1 Visão geral ───────────────────────────────────────────────────────────

def _consideracoes_alem_do_resumo(dados: dict) -> str:
    """A Prévia já usa o PRIMEIRO parágrafo das considerações como "Resumo do caso" —
    aqui entra só o que sobrou dele, para não repetir o mesmo texto duas vezes."""
    paragrafos = [p.strip() for p in _txt(dados.get("consideracoes_gerais")).split("\n") if p.strip()]
    return "\n".join(paragrafos[1:])


def _secao_visao_geral(doc, numero: str, dados: dict) -> bool:
    passivos = _pares_preenchidos([
        ("Passivo identificado — Fiscal", dados.get("passivo_fiscal")),
        ("Passivo identificado — Trabalhista", dados.get("passivo_trabalhista")),
        ("Passivo identificado — Cível", dados.get("passivo_civel")),
        ("Passivo identificado — Total", dados.get("passivo_total")),
    ])
    pares = _pares_preenchidos([
        ("Total atingível mapeado — VM", dados.get("total_atingivel_vm")),
        ("Total atingível mapeado — VF", dados.get("total_atingivel_vp")),
        ("Tese(s) principal(is)", _texto_analise(dados.get("teses_principais"))),
    ]) + passivos
    consideracoes = _consideracoes_alem_do_resumo(dados)
    if not pares and not consideracoes:
        return False

    _sub_orange(doc, f"{numero} Visão Geral do Caso")
    if pares:
        _kv_table(doc, ["INDICADOR", "VALOR / INFORMAÇÃO"], pares)
        _spacer(doc, pts=2)
    if consideracoes:
        _texto_livre(doc, "Principais Considerações", consideracoes)
    return True


# ── 6.2 Ativos consolidados ───────────────────────────────────────────────────

def _secao_ativos_consolidados(doc, numero: str, dados: dict) -> bool:
    resumo = _dicts(dados.get("visao_consolidada_ativos"))
    if not resumo:
        return False
    linhas = [[
        item.get("tese", ""),
        item.get("vm", ""),
        item.get("vp", ""),
        item.get("onus", ""),
        item.get("observacoes", ""),
    ] for item in resumo]
    total = [
        "TOTAL GERAL",
        _txt(dados.get("total_atingivel_vm")) or _somar_moeda(resumo, "vm"),
        _txt(dados.get("total_atingivel_vp")) or _somar_moeda(resumo, "vp"),
        _somar_moeda(resumo, "onus"),
        "",
    ]
    _sub_orange(doc, f"{numero} Visão Consolidada dos Ativos")
    _grid_table(doc, ["TESE", "VM (R$)", "VF (R$)", "ÔNUS (R$)", "OBSERVAÇÕES"],
                linhas, [3.53, 2.56, 2.56, 2.56, 5.31], total_row=total)
    _spacer(doc, pts=2)
    return True


# ── Redes sociais ─────────────────────────────────────────────────────────────

def _secao_redes_sociais(doc, numero: str, dados: dict) -> bool:
    """Perfis citados no material — rastro de padrão de vida e de patrimônio."""
    linhas = [[_txt(r.get("plataforma")), str(r.get("link") or "").strip(),
               _txt(r.get("observacoes"))]
              for r in _dicts(dados.get("redes_sociais"))]
    linhas = [linha for linha in linhas if linha[1]]
    if not linhas:
        return False
    _sub_orange(doc, f"{numero} Redes Sociais")
    _grid_table(doc, ["PLATAFORMA", "LINK", "OBSERVAÇÕES"], linhas, [3.4, 6.5, 7.0])
    _spacer(doc, pts=2)
    return True


# ── 6.3 Complemento da visão jurídica ─────────────────────────────────────────

def _lastros(doc, credito: dict) -> bool:
    """Um quadro por título executivo.

    Com um lastro só, instrumento e garantia já estão na linha correspondente do
    quadro da triagem e não se repetem aqui. Com mais de um, a triagem os traz
    numerados numa linha só — aí vale repetir, porque é o que separa um do outro.
    """
    itens = lastros_do_credito(credito)
    escreveu = False
    for ordem, lastro in enumerate(itens, 1):
        pares = [
            ("Data de Emissão", lastro.get("data_emissao")),
            ("Data do Vencimento", lastro.get("data_vencimento")),
            ("Partes", lastro.get("partes")),
            ("Destinação do Lastro", lastro.get("destinacao")),
            ("Assinaturas", lastro.get("assinaturas")),
        ]
        if len(itens) > 1:
            pares = ([("Lastro / Instrumento", lastro.get("lastro"))] + pares
                     + [("Garantia", lastro.get("garantia"))])
        linhas = _pares_preenchidos(pares)
        if not linhas:
            continue
        _sub_gray(doc, f"Lastro nº {ordem}" if len(itens) > 1 else "Lastro / Instrumento")
        _kv_label_table(doc, linhas)
        _spacer(doc, pts=2)
        escreveu = True
    return escreveu


def _complemento_do_credito(doc, credito: dict) -> bool:
    """Só o que a Prévia NÃO mostra no quadro "Dados do Processo" dela."""
    escreveu = _quadro_kv(doc, "Dados Complementares do Processo", [
        ("Vara / Comarca", credito.get("vara_comarca")),
        ("Critério de atualização do SAT", credito.get("criterio_sat")),
        ("Honorários", credito.get("honorarios")),
        ("Prescrição?", credito.get("prescricao")),
        ("Sucumbência?", credito.get("sucumbencia")),
        ("Riscos jurídicos gerais", _texto_analise(credito.get("riscos_juridicos"))),
    ])
    escreveu |= _lastros(doc, credito)
    escreveu |= _quadro_kv(doc, "Índices de Correção do Contrato — Adimplemento", [
        ("Correção monetária", credito.get("adimp_cm")),
        ("Juros remuneratórios", credito.get("adimp_jr")),
        ("Capitalização", credito.get("adimp_cap")),
    ])
    escreveu |= _quadro_kv(doc, "Índices de Correção do Contrato — Inadimplemento", [
        ("Correção monetária", credito.get("ind_cm")),
        ("Juros remuneratórios", credito.get("ind_jr")),
        ("Juros moratórios", credito.get("ind_jm")),
        ("Multa moratória", credito.get("ind_multa")),
        ("Capitalização", credito.get("ind_cap")),
        ("Comissão de permanência", credito.get("ind_comissao")),
    ])
    escreveu |= _quadro_kv(doc, "Planilha Inicial", [
        ("Correção monetária", credito.get("plan_cm")),
        ("Juros remuneratórios", credito.get("plan_jr")),
        ("Multa moratória", credito.get("plan_multa")),
        ("Capitalização", credito.get("plan_cap")),
        ("Comissão de permanência", credito.get("plan_comissao")),
        ("Ponderações", credito.get("plan_ponderacoes")),
    ])
    escreveu |= _quadro_kv(doc, "Última Memória de Cálculo", [
        ("Data da Juntada", credito.get("memoria_data_juntada")),
        ("Total Atualizado", credito.get("memoria_total")),
        ("Data-base", credito.get("memoria_data_base")),
        ("Índices aplicados", credito.get("memoria_indices")),
        ("Ponderações", credito.get("memoria_ponderacoes")),
    ])
    escreveu |= _quadro_grade(
        doc, "Citação", ["EXECUTADO", "MODALIDADE", "DATA", "FLS."], [6.0, 3.4, 3.5, 4.0],
        [[c.get("executado", ""), c.get("modalidade", ""), c.get("data", ""), c.get("fls", "")]
         for c in _dicts(credito.get("citacoes"))],
    )

    for ordem, embargo in enumerate(_dicts(credito.get("embargos")), 1):
        linhas = _pares_preenchidos([
            ("Embargante", embargo.get("embargante")),
            ("Data da Distribuição", embargo.get("data_dist")),
            ("Tese", embargo.get("tese")),
            ("Principais Andamentos", embargo.get("andamentos_resumo")),
            ("Status Atual", embargo.get("status")),
        ])
        if not linhas:
            continue
        _sub_gray(doc, f"{_txt(embargo.get('tipo')) or 'Embargos à Execução'} nº {ordem}")
        _kv_label_table(doc, linhas)
        _spacer(doc, pts=2)
        escreveu = True

    # Recurso: a Prévia já traz processo, polo ativo, finalidade e status — aqui só o resto.
    for ordem, recurso in enumerate(_dicts(credito.get("recursos")), 1):
        linhas = _pares_preenchidos([
            ("Decisão Recorrida", recurso.get("decisao_recorrida")),
            ("Data da Distribuição", recurso.get("data_dist")),
            ("Principais Andamentos", recurso.get("andamentos_resumo")),
        ])
        if not linhas:
            continue
        _sub_gray(doc, f"Recurso nº {ordem} — complemento")
        _kv_label_table(doc, linhas)
        _spacer(doc, pts=2)
        escreveu = True

    escreveu |= _quadro_grade(
        doc, "Principais Andamentos Processuais", ["DATA", "DESCRIÇÃO", "FLS. / EVENTO"],
        [2.6, 10.3, 4.0],
        [[a.get("data", ""), a.get("descricao", ""), a.get("fls", "")]
         for a in _dicts(credito.get("andamentos"))],
    )
    return escreveu


def _secao_visao_juridica(doc, numero: str, dados: dict) -> bool:
    creditos = _dicts(dados.get("creditos"))
    if not creditos:
        return False

    titulo = _sub_orange(doc, f"{numero} Complemento da Visão Jurídica")
    nota = _note(doc, "Um bloco por crédito, na mesma ordem e com o mesmo nome da seção 2.")
    escreveu = False
    for indice, credito in enumerate(creditos, 1):
        marcador = _para(doc, f"{numero}.{indice} {rotulo_credito(credito, indice)}",
                         bold=True, size=10, color=_TXT, before=8, after=3)
        if _complemento_do_credito(doc, credito):
            escreveu = True
        else:
            _remover(marcador)
    if not escreveu:
        _remover(titulo)
        _remover(nota)
    return escreveu


# ── 6.4 Teses de recuperação ──────────────────────────────────────────────────

def _tese_penhora(doc, dados, tese) -> bool:
    escreveu = _texto_livre(doc, "Ponderações e Observações Gerais", tese.get("analise"))
    return _quadro_ativos(doc, "Matrículas Mapeadas",
                          _filtrar_ativos(dados.get("ativos"), "Penhora Direta")) or escreveu


def _tese_idpj(doc, dados, tese) -> bool:
    empresa = _dict(tese.get("empresa_alvo"))
    escreveu = _texto_livre(doc, "Resumo da Tese", tese.get("resumo"))
    escreveu |= _quadro_kv(doc, "Empresa-Alvo", [
        ("Razão social da empresa-alvo", empresa.get("razao_social")),
        ("CNPJ", empresa.get("cnpj")),
        ("CNAE principal", empresa.get("cnae_principal")),
        ("Atividades secundárias", empresa.get("atividades_secundarias")),
        ("Sócio atual", empresa.get("socio_atual")),
        ("Endereço fiscal", empresa.get("endereco_fiscal")),
        ("Fundamentação da tese", empresa.get("fundamentacao")),
        ("Crédito mais adequado para propositura", empresa.get("credito_propositura")),
    ])
    escreveu |= _quadro_grade(
        doc, "Cronologia Societária — Atos Relevantes", ["DATA", "ATO", "DETALHAMENTO"],
        [2.6, 4.3, 10.0],
        [[item.get("data", ""), item.get("ato", ""),
          " — ".join(filter(None, [_txt(item.get("detalhamento")), _txt(item.get("referencia"))]))]
         for item in _dicts(tese.get("cronologia"))],
    )
    escreveu |= _texto_livre(doc, "Evidências de Controle Informal", tese.get("evidencias"))
    escreveu |= _quadro_ativos(doc, "Ativos Atingíveis via IDPJ",
                               _filtrar_ativos(dados.get("ativos"), "IDPJ"))
    escreveu |= _texto_livre(doc, "Upside Identificado", tese.get("upside"))
    return escreveu


def _tese_fraude(doc, dados, tese) -> bool:
    escreveu = _texto_livre(doc, "Resumo da Tese", tese.get("resumo"))
    escreveu |= _texto_livre(doc, "Má-fé e Insolvência", tese.get("ma_fe_insolvencia"))
    escreveu |= _quadro_ativos(doc, "Ativos Atingíveis via Fraude à Execução",
                               _filtrar_ativos(dados.get("ativos"), "Fraude à Execução"))
    return escreveu


def _secao_teses(doc, numero: str, dados: dict) -> bool:
    teses = _dict(dados.get("teses_recuperacao"))
    blocos = [
        ("Penhora Direta", _tese_penhora, _dict(teses.get("penhora_direta"))),
        ("IDPJ", _tese_idpj, _dict(teses.get("idpj"))),
        ("Fraude à Execução", _tese_fraude, _dict(teses.get("fraude_execucao"))),
    ]

    titulo = _sub_orange(doc, f"{numero} Teses de Recuperação")
    ordem = 0
    escreveu = False
    for nome, render, tese in blocos:
        if not tese:
            continue
        ordem += 1
        marcador = _para(doc, f"{numero}.{ordem} {nome}", bold=True, size=10,
                         color=_TXT, before=8, after=3)
        if render(doc, dados, tese):
            escreveu = True
        else:
            ordem -= 1
            _remover(marcador)

    outras = _texto_analise(teses.get("outras"))
    if outras:
        ordem += 1
        _para(doc, f"{numero}.{ordem} Outras teses", bold=True, size=10,
              color=_TXT, before=8, after=3)
        for paragrafo in [p for p in outras.split("\n") if p.strip()]:
            _body(doc, paragrafo)
        _spacer(doc, pts=2)
        escreveu = True

    if not escreveu:
        _remover(titulo)
    return escreveu


# ── 6.5 Quadros pedidos na instrução adicional ────────────────────────────────

def _secao_quadros_extras(doc, numero: str, dados: dict) -> bool:
    quadros = []
    for quadro in _dicts(dados.get("quadros_extras")):
        colunas = [_txt(c) for c in (quadro.get("colunas") or [])]
        linhas = [[_txt(celula) for celula in linha][:len(colunas)]
                  for linha in (quadro.get("linhas") or []) if isinstance(linha, (list, tuple))]
        linhas = [linha for linha in linhas if any(linha)]
        if colunas and linhas:
            quadros.append((_txt(quadro.get("titulo")) or "Quadro complementar", colunas, linhas))
    if not quadros:
        return False

    _sub_orange(doc, f"{numero} Quadros Complementares")
    for titulo, colunas, linhas in quadros:
        largura = round(16.9 / len(colunas), 2)
        _sub_gray(doc, titulo)
        _grid_table(doc, colunas, linhas, [largura] * len(colunas))
        _spacer(doc, pts=2)
    return True


# ── Montagem ──────────────────────────────────────────────────────────────────

def _remover(paragrafo) -> None:
    """Tira do corpo um título que acabou sem conteúdo embaixo."""
    elemento = paragrafo._p
    if elemento.getparent() is not None:
        elemento.getparent().remove(elemento)


def montar_capitulo_opcional(doc, dados: dict) -> bool:
    """Acrescenta o capítulo "6. OPCIONAL" ao fim da Prévia.

    Devolve False (e não escreve nada) quando a análise não preencheu nenhum dos
    campos que só existem no PPA — caso em que o capítulo seria só uma casca de
    títulos vazios.
    """
    dados = dados if isinstance(dados, dict) else {}

    # Sem normalizar: os títulos de capítulo do template da Prévia são em caixa alta,
    # e a normalização de nomes próprios devolveria "6. Opcional".
    titulo = _para(doc, TITULO_CAPITULO, bold=True, size=13, color=_LARANJA,
                   before=14, after=5, normalizar=False)
    nota = _note(
        doc,
        "Conteúdo do Dossiê PPA que não cabe nos quadros da Prévia. Só aparece o que a "
        "análise efetivamente preencheu — quadros que dependem de e-CAC, certidões, "
        "avaliações e demais diligências externas ao processo continuam no Dossiê PPA.",
    )

    secoes = [
        _secao_visao_geral,
        _secao_ativos_consolidados,
        _secao_redes_sociais,
        _secao_visao_juridica,
        _secao_teses,
        _secao_quadros_extras,
    ]
    ordem = 0
    for secao in secoes:
        proximo = f"6.{ordem + 1}"
        if secao(doc, proximo, dados):
            ordem += 1

    if ordem == 0:
        _remover(titulo)
        _remover(nota)
        return False
    return True
