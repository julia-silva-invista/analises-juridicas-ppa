# -*- coding: utf-8 -*-
"""
Análises Jurídicas — Invista
Robô consolidado: Análise de Processos · Recuperação Judicial · Matrículas
"""
import os
import json
import inspect
import tempfile

import gradio as gr

from design import CSS, HEADER_HTML, FOOTER_HTML
from processos import (
    proc_analisar, proc_gerar_word, proc_gerar_dossie, proc_gerar_previa,
    proc_gerar_cronologia, proc_responder,
)
from cronologia_prescricao import (
    cronologia_aplicar_html,
    exportar_html as presc_exportar_html,
)
from rj import rj_analisar, rj_gerar_word, rj_responder, rj_gerar_excel_credores, rj_gerar_checklist, rj_gerar_checklist_creditos
from matriculas import mat_gerar_excel, mat_responder
from coleta import coleta_gerar, coleta_gerar_dossie_dispatch
from analysis_runtime import environment_status_json
from timeline_societaria import (
    timeline_analisar,
    timeline_gerar_word,
    timeline_exportar_html,
    timeline_aplicar_html,
)

os.makedirs("resultados", exist_ok=True)
os.makedirs("tmp_pdfs", exist_ok=True)

# Garante que o diretorio de cache de upload do Gradio exista ANTES do primeiro upload --
# em containers recem-reiniciados, o diretorio padrao (<tmp>/gradio) pode nao existir ainda
# quando a primeira requisicao chega, causando FileNotFoundError no preprocess do gr.File.
_GRADIO_TEMP_DIR = os.environ.setdefault("GRADIO_TEMP_DIR", os.path.join(tempfile.gettempdir(), "gradio"))
os.makedirs(_GRADIO_TEMP_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE GRADIO
# ══════════════════════════════════════════════════════════════════════════════

gr.close_all()

_FORCE_LIGHT_JS = """
() => {
    const url = new URL(window.location.href);
    if (url.searchParams.get('__theme') !== 'light') {
        url.searchParams.set('__theme', 'light');
        window.location.replace(url.toString());
        return;
    }
    document.documentElement.classList.remove('dark');
    document.documentElement.setAttribute('data-theme', 'light');
    const observer = new MutationObserver(() => {
        if (document.documentElement.classList.contains('dark')) {
            document.documentElement.classList.remove('dark');
        }
    });
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });

    // Pilula de status no cabecalho conforme a demanda real de analises.
    const applyStatus = () => {
        const card = document.querySelector(".inv-status-card");
        const value = document.querySelector(".inv-status-value");
        const field = document.querySelector("#environment-status-json textarea, #environment-status-json input");
        if (!card || !value || !field) return;
        try {
            const status = JSON.parse(field.value || "{}");
            const nextState = status.state || "stable";
            const nextLabel = status.label || "Estável";
            const nextTitle = `${status.active || 0} análise(s) ativa(s) · ${status.waiting || 0} na fila`;
            if (card.dataset.status !== nextState) card.dataset.status = nextState;
            if (value.textContent !== nextLabel) value.textContent = nextLabel;
            if (card.title !== nextTitle) card.title = nextTitle;
        } catch (_) {}
    };
    applyStatus();
    window.setInterval(applyStatus, 1000);

    // Carrossel de abas — setinhas discretas quando as abas nao cabem na largura da tela.
    const setupTabCarousel = () => {
        const tabNav = document.querySelector(".tab-nav [role='tablist']");
        if (!tabNav || tabNav.dataset.carouselReady) return;
        tabNav.dataset.carouselReady = "1";

        const wrapper = tabNav.parentElement;
        if (wrapper && getComputedStyle(wrapper).position === "static") {
            wrapper.style.position = "relative";
        }

        const prevBtn = document.createElement("button");
        prevBtn.type = "button";
        prevBtn.className = "tab-nav-arrow tab-nav-arrow--prev";
        prevBtn.setAttribute("aria-label", "Ver abas anteriores");
        prevBtn.textContent = "‹";

        const nextBtn = document.createElement("button");
        nextBtn.type = "button";
        nextBtn.className = "tab-nav-arrow tab-nav-arrow--next";
        nextBtn.setAttribute("aria-label", "Ver mais abas");
        nextBtn.textContent = "›";

        wrapper.appendChild(prevBtn);
        wrapper.appendChild(nextBtn);

        const updateArrows = () => {
            const maxScroll = tabNav.scrollWidth - tabNav.clientWidth;
            prevBtn.classList.toggle("is-visible", tabNav.scrollLeft > 4);
            nextBtn.classList.toggle("is-visible", tabNav.scrollLeft < maxScroll - 4);
        };

        prevBtn.addEventListener("click", () => tabNav.scrollBy({ left: -180, behavior: "smooth" }));
        nextBtn.addEventListener("click", () => tabNav.scrollBy({ left: 180, behavior: "smooth" }));
        tabNav.addEventListener("scroll", updateArrows);
        window.addEventListener("resize", updateArrows);
        updateArrows();
    };
    setupTabCarousel();
    new MutationObserver(setupTabCarousel).observe(document.body, { childList: true, subtree: true });
}
"""

def _editar_no_html_js(seletor_raiz: str) -> str:
    """Liga/desliga a edição no próprio painel e serializa o resultado.

    Mesmo modelo do Painel de Deals: `contenteditable` nos campos marcados com
    data-tl-campo, contorno tracejado, e o botão troca de rótulo. Ao concluir, percorre
    a árvore e devolve JSON — os itens são lidos por POSIÇÃO dentro de cada bloco, então
    acrescentar e remover linha funciona sem renumerar nada.

    Recebe (modo_ligado, ponte) e devolve [novo_modo, json] — evento único com js= e fn=,
    porque `.then` encadeado após evento só-JS não dispara no Gradio 5.29.1 do Space.
    """
    return """
async (ligado, _ponte) => {
    const raiz = document.querySelector("__RAIZ__");
    if (!raiz) return [String(ligado), ""];

    const editaveis = () => raiz.querySelectorAll("[data-tl-campo]");
    // textContent e nao innerText: innerText aplica o text-transform do CSS, e o
    // cabecalho da coluna e maiusculo por estilo — editar "7a Alteracao" ali gravaria
    // "7A ALTERACAO" no JSON.
    const texto = (el) => (el.textContent || "").replace(/\\s+/g, " ").trim();

    if (String(ligado) !== "1") {
        raiz.classList.add("tl2-edit-mode");
        editaveis().forEach((el) => { el.contentEditable = "true"; });

        if (!raiz.dataset.tlLigado) {
            raiz.dataset.tlLigado = "1";
            // Data e ato aparecem duas vezes na mesma coluna: no cabecalho e no pop-up
            // de detalhamento. A serializacao le a primeira ocorrencia, entao editar a
            // do pop-up sumiria sem aviso. Espelhar enquanto se digita mantem as duas
            // iguais e faz a duplicata parar de importar.
            raiz.addEventListener("input", (ev) => {
                const campo = ev.target && ev.target.closest
                    ? ev.target.closest("[data-tl-campo]") : null;
                if (!campo) return;
                const coluna = campo.closest("[data-tl-evento]");
                if (!coluna || campo.closest("[data-tl-lista]")) return;
                const nome = campo.getAttribute("data-tl-campo");
                coluna.querySelectorAll('[data-tl-campo="' + nome + '"]').forEach((par) => {
                    if (par !== campo && !par.closest("[data-tl-lista]")
                        && par.textContent !== campo.textContent) {
                        par.textContent = campo.textContent;
                    }
                });
            });
            raiz.addEventListener("click", (ev) => {
                const alvo = ev.target;
                if (!alvo || !alvo.matches) return;
                if (alvo.matches("[data-tl-remover]")) {
                    ev.preventDefault();
                    const linha = alvo.closest("[data-tl-item]");
                    if (linha) linha.remove();
                } else if (alvo.matches("[data-tl-adicionar-ato]")) {
                    ev.preventDefault();
                    const trilha = raiz.querySelector(".tl2-track");
                    const modelo = raiz.querySelector("[data-tl-evento]");
                    if (!trilha || !modelo) return;
                    const novo = modelo.cloneNode(true);
                    novo.setAttribute("data-tl-extra", "{}");
                    novo.querySelectorAll("[data-tl-item]").forEach((n) => n.remove());
                    novo.querySelectorAll("[data-tl-campo]").forEach((n) => { n.textContent = ""; });
                    // Cada modal tem id proprio; sem renumerar, dois atos abririam o mesmo.
                    const sufixo = "novo-" + Date.now();
                    const marca = novo.querySelector(".tl2-toggle");
                    if (marca) marca.id = "tl2-det-" + sufixo;
                    novo.querySelectorAll("label[for^='tl2-det-']").forEach(
                        (l) => l.setAttribute("for", "tl2-det-" + sufixo));
                    trilha.appendChild(novo);
                    trilha.style.setProperty("--tl2-cols", trilha.children.length);
                } else if (alvo.matches("[data-tl-remover-ato]")) {
                    ev.preventDefault();
                    const coluna = alvo.closest("[data-tl-evento]");
                    const trilha = coluna && coluna.parentElement;
                    if (coluna) coluna.remove();
                    if (trilha) trilha.style.setProperty("--tl2-cols", trilha.children.length);
                } else if (alvo.matches("[data-tl-adicionar]")) {
                    ev.preventDefault();
                    const bloco = alvo.closest("[data-tl-lista]");
                    // O molde da linha em branco vem do servidor, em data-tl-item-html.
                    // Remontar a linha aqui era manter uma segunda versao da estrutura,
                    // e foi por isso que o "+ caixa" nao inseria nada: a caixinha do ato
                    // e um <div> solto no bloco, nao um <li> dentro de <ul>.
                    const molde = bloco ? bloco.getAttribute("data-tl-item-html") : "";
                    if (!molde) return;
                    const provisorio = document.createElement("div");
                    provisorio.innerHTML = molde;
                    const item = provisorio.firstElementChild;
                    if (!item) return;
                    // Lista com <ul> propria: entra no fim da <ul>. Bloco sem <ul> (as
                    // caixinhas): entra antes do proprio botao, que e filho do bloco.
                    const lista = bloco.querySelector(".tl2-list") || bloco;
                    lista.insertBefore(item, alvo.parentElement === lista ? alvo : null);
                    item.querySelectorAll("[data-tl-campo]").forEach(
                        (c) => { c.contentEditable = "true"; });
                    const primeiro = item.querySelector("[data-tl-campo]");
                    if (primeiro) primeiro.focus();
                }
            });
        }
        return ["1", ""];
    }

    // Concluindo: lê a árvore inteira do DOM e devolve para o servidor.
    const arvore = {eventos: []};
    // Qualquer campo do cabecalho da secao — empresa, cnpj, titulo do lastro, numero do
    // processo. Ler por nome, e nao um a um, evita esquecer de ligar o proximo.
    raiz.querySelectorAll("[data-tl-campo]").forEach((el) => {
        if (el.closest("[data-tl-evento]")) return;
        const nome = el.getAttribute("data-tl-campo");
        if (nome && !(nome in arvore)) arvore[nome] = texto(el);
    });
    try { arvore._extra = JSON.parse(raiz.getAttribute("data-tl-extra") || "{}"); }
    catch (e) { arvore._extra = {}; }

    raiz.querySelectorAll("[data-tl-evento]").forEach((col) => {
        const evento = {};
        try { evento._extra = JSON.parse(col.getAttribute("data-tl-extra") || "{}"); }
        catch (e) { evento._extra = {}; }
        col.querySelectorAll("[data-tl-campo]").forEach((el) => {
            // Campo de item de lista é lido no laço de baixo, não aqui.
            if (el.closest("[data-tl-lista]")) return;
            const nome = el.getAttribute("data-tl-campo");
            if (nome && !(nome in evento)) evento[nome] = texto(el);
        });
        col.querySelectorAll("[data-tl-lista]").forEach((bloco) => {
            const chave = bloco.getAttribute("data-tl-lista");
            evento[chave] = [...bloco.querySelectorAll("[data-tl-item]")].map((linha) => {
                const item = {};
                linha.querySelectorAll("[data-tl-campo]").forEach((cel) => {
                    item[cel.getAttribute("data-tl-campo")] = texto(cel);
                });
                return item;
            });
        });
        arvore.eventos.push(evento);
    });

    raiz.classList.remove("tl2-edit-mode");
    editaveis().forEach((el) => { el.contentEditable = "false"; });
    return ["0", JSON.stringify(arvore)];
}
""".replace("__RAIZ__", seletor_raiz)


_EDITAR_TIMELINE_JS = _editar_no_html_js("#timeline-export-area")
_EDITAR_CRONOLOGIA_JS = _editar_no_html_js("#cronologia-prescricao-area")

with gr.Blocks(
    title="Análises Jurídicas — Invista",
    css=CSS,
    js=_FORCE_LIGHT_JS,
    theme=gr.themes.Default(),
) as demo:
    gr.HTML(HEADER_HTML)
    environment_status_state = gr.Textbox(
        value=environment_status_json(), container=False, elem_id="environment-status-json"
    )
    environment_status_timer = gr.Timer(value=5.0, active=True)

    with gr.Tabs(elem_classes=["tab-nav"]):

        # ── Tab 1: Análise de Processos ──────────────────────────────────────
        with gr.Tab("Processos"):
            with gr.Row(elem_classes=["input-grid-three"]):
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    proc_pdf_principal = gr.File(
                        label="Processo principal",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    proc_pdf_relacionados = gr.File(
                        label="Processos relacionados (opcional)",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    with gr.Group(elem_classes=["instructions-with-gemini"]):
                        proc_instrucoes = gr.Textbox(
                            label="Instruções adicionais",
                            placeholder="Ex: verificar penhora sobre imóvel; analisar menção a prescrição; checar bem de família...",
                            lines=4,
                            elem_classes=["instructions-field"],
                        )
                        proc_versao_resumida = gr.Checkbox(
                            label="Versão resumida",
                            value=False,
                            elem_classes=["cb-slot", "cb-tip-resumida"],
                        )

            with gr.Row(elem_classes=["analysis-action-row"]):
                proc_analisar_btn = gr.Button("Analisar processo", variant="primary", elem_classes=["analysis-run-btn"])

            with gr.Tabs():
                with gr.Tab("Progresso"):
                    proc_log = gr.Textbox(
                        label="Log de execução",
                        lines=6,
                        interactive=False,
                        elem_classes=["log-area"],
                        placeholder="O andamento da análise aparecerá aqui em tempo real...",
                    )
                with gr.Tab("Relatório"):
                    proc_report = gr.Textbox(
                        label="Relatório Jurídico",
                        lines=50,
                        interactive=False,
                        placeholder="O relatório aparecerá aqui após a conclusão da análise...",
                    )

            proc_relatorio_state = gr.State("")
            proc_extracao_state  = gr.State("")   # texto OCR completo (fonte para os dossiês)

            with gr.Row():
                proc_word_btn   = gr.Button("Baixar Word",        variant="secondary", elem_classes=["word-download-btn"])
                proc_previa_btn = gr.Button("Dossiê Prévia",      variant="secondary", elem_classes=["word-download-btn"])
                proc_dossie_btn = gr.Button("Dossiê Desalinhado", variant="secondary", elem_classes=["word-download-btn"])
            proc_word_file = gr.File(
                label="", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            proc_previa_file = gr.File(
                label="Dossiê Prévia", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            proc_previa_status = gr.Markdown("")
            proc_dossie_file = gr.File(
                label="Dossiê Desalinhado", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            proc_dossie_status = gr.Markdown("")

            # Cronologia da prescrição: o modelo extrai os marcos datados; o regime de
            # cada intervalo e as viradas de lei são calculados em Python.
            with gr.Row():
                proc_cronologia_btn = gr.Button(
                    "Cronologia processual — prescrição intercorrente",
                    variant="secondary", elem_classes=["word-download-btn"],
                )
            proc_presc_state = gr.State({})
            proc_presc_html = gr.HTML()
            # Igual à timeline societária: editar e exportar só aparecem quando existe
            # cronologia na tela para editar e exportar.
            with gr.Row(visible=False) as proc_presc_acoes_row:
                proc_presc_editar_btn = gr.Button("Editar cronologia", variant="secondary")
                proc_presc_export_btn = gr.DownloadButton("Exportar cronologia (HTML)",
                                                          variant="secondary")
            proc_presc_editando = gr.Textbox(value="0", visible=False)
            proc_presc_ponte = gr.Textbox(visible=False)

            gr.HTML('<hr class="inv-divider">')
            with gr.Column(elem_classes=["qa-section"]):
                with gr.Row(elem_classes=["qa-ask-row"]):
                    with gr.Column(scale=4):
                        proc_pergunta = gr.Textbox(
                            label="Perguntas sobre a análise:",
                            lines=3,
                            placeholder="Exemplos: Quando foi realizado o primeiro pedido de penhora? Há risco de prescrição intercorrente? Há risco de sucumbência? Qual o valor atualizado da dívida?"
                        )
                    with gr.Column(scale=1, elem_classes=["qa-ask-col"]):
                        proc_perguntar_btn = gr.Button("Perguntar", variant="primary")
                proc_resposta = gr.Textbox(label="Resposta", lines=5, interactive=False)

        # ── Tab 2: Análise de Recuperação Judicial ───────────────────────────
        with gr.Tab("Recuperação Judicial"):
            RJ_MAX_CRED = 12
            rj_cred_count = gr.State(1)
            rj_cred_rows, rj_cred_nomes, rj_cred_docs = [], [], []

            with gr.Row(elem_classes=["input-grid-three"]):
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    rj_pdf_principal = gr.File(
                        label="Processo de RJ",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    rj_pdf_relacionados = gr.File(
                        label="Processos relacionados (opcional)",
                        file_types=[".pdf", ".PDF", ".docx"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    gr.Markdown(
                        "**Como usar:**\n\n"
                        "1. Envie o **processo de RJ** para gerar o **Checklist de RJ**.\n"
                        "2. Envie os **processos relacionados** (execuções, impugnações de crédito) — e a "
                        "RJ, opcionalmente — para gerar o **Checklist de Créditos** de um crédito "
                        "específico.\n"
                        "3. Havendo mais de um credor, recomendamos listá-los em **Checklist de Créditos "
                        "— Inserir dados do credor** (abaixo): o robô gera um checklist de crédito para "
                        "cada credor informado.\n"
                        "4. Já tem uma análise de RJ pronta? Envie o **Word gerado** (em vez do PDF) na "
                        "caixa de **processos relacionados**, junto com os processos relacionados novos — "
                        "gera o Checklist de Créditos sem repetir a extração inteira da RJ."
                    )

            with gr.Accordion("Checklist de Créditos — Inserir dados do credor (opcional)", open=False, elem_classes=["mat-accordion"]):
                gr.Markdown(
                    "Informe o nome e o CPF/CNPJ de cada credor-alvo (use **+ Adicionar credor** para "
                    "incluir mais). Deixe em branco para a IA identificar o crédito automaticamente "
                    "(a partir do polo ativo das execuções relacionadas)."
                )
                for _i in range(RJ_MAX_CRED):
                    with gr.Row(visible=(_i == 0), elem_classes=["mat-accordion-row"]) as _crow:
                        _cnome = gr.Textbox(label="Nome do credor", scale=3, container=True,
                                            placeholder="Ex: BASF S.A.")
                        _cdoc  = gr.Textbox(label="CPF/CNPJ", scale=2, container=True,
                                            placeholder="00.000.000/0001-00")
                    rj_cred_rows.append(_crow)
                    rj_cred_nomes.append(_cnome)
                    rj_cred_docs.append(_cdoc)
                rj_cred_add_btn = gr.Button("+ Adicionar credor", variant="secondary")

            with gr.Row(elem_classes=["analysis-action-row"]):
                rj_analisar_btn = gr.Button("Analisar", variant="primary", elem_classes=["analysis-run-btn"])

            with gr.Tabs():
                with gr.Tab("Progresso"):
                    rj_log = gr.Textbox(
                        label="Log de execução",
                        lines=6,
                        interactive=False,
                        elem_classes=["log-area"],
                        placeholder="O andamento da análise aparecerá aqui em tempo real...",
                    )
                with gr.Tab("Relatório"):
                    rj_report = gr.Textbox(
                        label="Relatório — Recuperação Judicial",
                        lines=50,
                        interactive=False,
                        placeholder="O relatório aparecerá aqui após a conclusão da análise...",
                    )

            rj_relatorio_state = gr.State("")
            rj_extracao_state  = gr.State("")   # texto bruto das extrações (fonte para o Excel de credores)

            with gr.Row():
                rj_word_btn           = gr.Button("Baixar Word",             variant="secondary", elem_classes=["word-download-btn"])
                rj_excel_cred_btn     = gr.Button("Gerar Excel de Credores", variant="secondary", elem_classes=["word-download-btn"])
                rj_checklist_btn      = gr.Button("Checklist RJ",            variant="secondary", elem_classes=["word-download-btn"])
                rj_checklist_cred_btn = gr.Button("Gerar Checklist de Créditos", variant="secondary", elem_classes=["word-download-btn"])
            rj_word_file = gr.File(
                label="", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            rj_excel_cred_file = gr.File(
                label="Excel de Credores", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            rj_excel_cred_status = gr.Textbox(label="", interactive=False, lines=1, show_label=False)
            rj_checklist_file = gr.File(
                label="Checklist RJ", interactive=False, visible=False, height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            rj_checklist_status    = gr.Markdown("")
            rj_checklist_cred_file = gr.File(
                label="Checklist de Créditos", interactive=False, visible=False,
                file_count="multiple", height=72,
                elem_classes=["word-file-output", "compact-file-output"],
            )
            rj_checklist_cred_status = gr.Markdown("")

            gr.HTML('<hr class="inv-divider">')
            with gr.Column(elem_classes=["qa-section"]):
                with gr.Row(elem_classes=["qa-ask-row"]):
                    with gr.Column(scale=4):
                        rj_pergunta = gr.Textbox(
                            label="Perguntas sobre a análise:",
                            lines=3,
                            placeholder="Exemplos: Qual é o maior credor e qual percentual detém dos créditos? Qual o status atual do stay period? Há risco de prescrição intercorrente nas execuções?"
                        )
                    with gr.Column(scale=1, elem_classes=["qa-ask-col"]):
                        rj_perguntar_btn = gr.Button("Perguntar", variant="primary")
                rj_resposta = gr.Textbox(label="Resposta", lines=5, interactive=False)

        # ── Tab 3: Análise de Matrículas ─────────────────────────────────────
        MAT_MAX_PESSOAS = 12

        with gr.Tab("Matrículas"):
            with gr.Row():
                with gr.Column(scale=2):
                    mat_arquivos = gr.File(
                        label="PDFs das Matrículas",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1):
                    gr.Markdown(
                        "**Como usar:**\n\n"
                        "1. Faça upload das matrículas (uma ou mais)\n"
                        "2. Preencha os dados da execução (opcional) para colorir alertas\n"
                        "3. Clique em **Gerar Excel**\n"
                        "4. Baixe o arquivo gerado\n\n"
                        "_Suporta matrículas escaneadas ou digitais._"
                    )

            with gr.Accordion("Inserir dados da execução (opcional)", open=False, elem_classes=["mat-accordion"]):
                with gr.Row():
                    mat_data_ajuizamento = gr.Textbox(
                        label="Data do ajuizamento",
                        placeholder="DD/MM/AAAA",
                        lines=1,
                        scale=1,
                    )
                # Mesmo padrão da aba de RJ: MAX linhas criadas de uma vez, só a
                # primeira visível, e o botão apenas revela a próxima. Antes eram dois
                # textareas livres e o NOME era descartado no parse — quem digitasse só
                # o nome não gerava alerta nenhum, em silêncio.
                mat_dev_count = gr.State(1)
                mat_grp_count = gr.State(1)
                mat_dev_rows, mat_dev_nomes, mat_dev_docs = [], [], []
                mat_grp_rows, mat_grp_nomes, mat_grp_docs = [], [], []

                def _mat_bloco_pessoas(titulo, placeholder_nome, linhas, nomes, docs):
                    gr.Markdown(f"**{titulo}**")
                    for indice in range(MAT_MAX_PESSOAS):
                        with gr.Row(visible=(indice == 0), elem_classes=["mat-accordion-row"]) as linha:
                            campo_nome = gr.Textbox(
                                label="Nome", scale=3, container=True,
                                placeholder=placeholder_nome,
                            )
                            campo_doc = gr.Textbox(
                                label="CPF/CNPJ", scale=2, container=True,
                                placeholder="000.000.000-00",
                            )
                        linhas.append(linha)
                        nomes.append(campo_nome)
                        docs.append(campo_doc)

                _mat_bloco_pessoas("Inserir devedor", "Ex: João da Silva",
                                   mat_dev_rows, mat_dev_nomes, mat_dev_docs)
                mat_dev_add_btn = gr.Button("+ Adicionar mais uma pessoa", variant="secondary")

                _mat_bloco_pessoas("Inserir pessoa do grupo", "Ex: Holdings XYZ S/A",
                                   mat_grp_rows, mat_grp_nomes, mat_grp_docs)
                mat_grp_add_btn = gr.Button("+ Adicionar mais uma pessoa", variant="secondary")

                gr.Markdown(
                    "🟡 **Amarelo claro** — transmissão envolvendo devedor ou pessoa do grupo  \n"
                    "🔴 **Vermelho claro** — alienação **pelo devedor** após o ajuizamento "
                    "(fraude à execução)  \n"
                    "_Basta o nome: sem CPF/CNPJ, o cruzamento é feito pelo nome._"
                )

            with gr.Row(elem_classes=["analysis-action-row"]):
                mat_gerar_btn = gr.Button("Gerar Excel", variant="primary", size="lg", elem_classes=["analysis-run-btn"])

            with gr.Tabs():
                with gr.Tab("Progresso"):
                    mat_log = gr.Textbox(
                        label="Status",
                        lines=10,
                        interactive=False,
                        elem_classes=["log-area"],
                        placeholder="O andamento aparecerá aqui...",
                    )
                with gr.Tab("Download"):
                    mat_status = gr.Textbox(label="Status", interactive=False)
                    mat_excel = gr.File(
                        label="Excel gerado", interactive=False, height=72,
                        elem_classes=["compact-file-output"],
                    )

            gr.HTML('<hr class="inv-divider">')
            with gr.Column(elem_classes=["qa-section"]):
                with gr.Row(elem_classes=["qa-ask-row"]):
                    with gr.Column(scale=4):
                        mat_pergunta = gr.Textbox(
                            label="Perguntas sobre as matrículas:",
                            lines=3,
                            placeholder="Exemplos: Indique as matrículas que pertencem a determinado devedor. Quais têm penhora vigente? Qual o valor total dos ônus?"
                        )
                    with gr.Column(scale=1, elem_classes=["qa-ask-col"]):
                        mat_perguntar_btn = gr.Button("Perguntar", variant="primary")
                mat_resposta = gr.Textbox(label="Resposta", lines=5, interactive=False)

        # ── Tab 4: Timeline Societária ────────────────────────────────────────
        with gr.Tab("Timeline Societária"):
            with gr.Row():
                with gr.Column(scale=2):
                    tl_arquivos = gr.File(
                        label="Atos societários (PDFs)",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1):
                    gr.Markdown(
                        "**Como usar:**\n\n"
                        "1. Faça upload de todos os atos societários (contrato social, "
                        "alterações, ACS)\n"
                        "2. Clique em **Analisar timeline societária**\n"
                        "3. Edite os eventos se necessário — ao concluir, os atos são "
                        "reordenados por data\n"
                        "4. Exporte como HTML interativo ou tabela em Word\n\n"
                        "_Suporta documentos escaneados (OCR)._"
                    )

            with gr.Row(elem_classes=["analysis-action-row"]):
                tl_analisar_btn = gr.Button(
                    "Analisar timeline societária", variant="primary", elem_classes=["analysis-run-btn"]
                )

            tl_status = gr.Textbox(
                label="Status",
                lines=2,
                interactive=False,
                elem_classes=["log-area"],
                placeholder="O andamento da análise aparecerá aqui...",
            )
            tl_timeline_html = gr.HTML()
            tl_data_state = gr.State({})

            # Edição no próprio painel, como no Painel de Deals: o botão liga
            # contenteditable nos campos e "Concluir edição" devolve a árvore.
            tl_editando = gr.Textbox(value="0", visible=False)
            tl_edicao_ponte = gr.Textbox(visible=False)

            # Editar e exportar só existem depois que há timeline na tela: antes disso os
            # botões não têm sobre o que agir, e a linha vazia só ocupava espaço.
            with gr.Row(visible=False) as tl_acoes_row:
                tl_editar_btn = gr.Button("Editar timeline", variant="secondary")
                tl_exportar_html_btn = gr.DownloadButton("Exportar HTML", variant="secondary")
                tl_gerar_tabela_btn = gr.Button("Gerar tabela (Word)", variant="secondary")

            # A caixa de download aparece com o arquivo, não antes dele.
            tl_tabela_word_file = gr.File(
                label="Tabela em Word", interactive=False, visible=False, height=72,
                elem_classes=["compact-file-output"],
            )

        # ── Tab 5: Coleta de Informações ─────────────────────────────────────
        with gr.Tab("Coleta de Informações"):
            with gr.Row():
                with gr.Column(scale=1):
                    coleta_excel_in = gr.File(
                        label="Excel(s) da Predictus",
                        file_types=[".xlsx", ".xls"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1):
                    coleta_excel_coleta_in = gr.File(
                        label="Excel da Coleta de Informações (Matrículas/Fiscal & Cível/Trabalhista)",
                        file_types=[".xlsx", ".xls"],
                    )
                with gr.Column(scale=1):
                    coleta_dossie_in = gr.File(
                        label="Dossiê PPA em Word (opcional — para atualizar)",
                        file_types=[".docx"],
                    )
                with gr.Column(scale=1):
                    gr.Markdown(
                        "**Como usar:**\n\n"
                        "1. Exporte o dossiê jurídico da Predictus em Excel\n"
                        "2. Faça upload de um ou mais arquivos\n"
                        "3. **Gerar Planilha** → planilha consolidada\n"
                        "4. Ou envie também o dossiê (Word) e clique em **Gerar Dossiê Atualizado** "
                        "para preencher a Seção 3 (Passivo)\n\n"
                        "5. Já tem o **Excel da Coleta de Informações** (abas Matrículas/Fiscal & "
                        "Cível/Trabalhista)? Envie-o junto com o Dossiê PPA e clique em **Gerar "
                        "Dossiê Atualizado** — preenche também os Ativos Atingíveis (agrupados por "
                        "Tese) além do Passivo.\n\n"
                        "_Suporta múltiplos devedores simultaneamente._"
                    )

            with gr.Row():
                coleta_gerar_btn        = gr.Button("Gerar Planilha", variant="primary", size="lg")
                coleta_gerar_dossie_btn = gr.Button("Gerar Dossiê Atualizado", variant="secondary", size="lg")

            with gr.Tabs():
                with gr.Tab("Progresso"):
                    coleta_log = gr.Textbox(
                        label="Log",
                        lines=12,
                        interactive=False,
                        elem_classes=["log-area"],
                        placeholder="O andamento aparecerá aqui...",
                    )
                with gr.Tab("Download"):
                    coleta_status = gr.Textbox(label="Status", interactive=False)
                    coleta_excel_out = gr.File(
                        label="Planilha preenchida", interactive=False, height=72,
                        elem_classes=["compact-file-output"],
                    )
                    coleta_dossie_out = gr.File(
                        label="Dossiê atualizado (passivo)", interactive=False, height=72,
                        elem_classes=["compact-file-output"],
                    )

    gr.HTML(FOOTER_HTML)

    # ── Eventos ──────────────────────────────────────────────────────────────

    # Processos — análise apenas ao clicar no botão
    proc_analisar_btn.click(
        fn=proc_analisar,
        inputs=[proc_pdf_principal, proc_pdf_relacionados, proc_instrucoes, proc_versao_resumida],
        outputs=[proc_log, proc_report, proc_relatorio_state, proc_extracao_state],
        concurrency_limit=20,
        concurrency_id="analises_pdf",
    )
    proc_word_btn.click(fn=proc_gerar_word, inputs=[proc_relatorio_state], outputs=[proc_word_file])
    # A instrução adicional vai junto: é o que permite pedir um campo ou uma tabela
    # que o template do dossiê não prevê.
    proc_previa_btn.click(fn=proc_gerar_previa, inputs=[proc_relatorio_state, proc_extracao_state, proc_instrucoes], outputs=[proc_previa_file, proc_previa_status])
    proc_dossie_btn.click(fn=proc_gerar_dossie, inputs=[proc_relatorio_state, proc_extracao_state, proc_instrucoes], outputs=[proc_dossie_file, proc_dossie_status])

    proc_cronologia_btn.click(
        fn=proc_gerar_cronologia,
        inputs=[proc_relatorio_state, proc_extracao_state],
        outputs=[proc_presc_state, proc_presc_html, proc_presc_acoes_row],
        concurrency_limit=2,
    )
    # Edição no próprio painel, igual à da timeline: o JS liga o contenteditable e
    # devolve a árvore; o servidor refaz o enquadramento e o cálculo.
    proc_presc_editar_btn.click(
        fn=cronologia_aplicar_html,
        # Sem gr.State atravessando o JS: o que o painel não desenha
        # (vencimento do título, nº do processo) viaja em data-tl-extra na raiz.
        inputs=[proc_presc_editando, proc_presc_ponte],
        outputs=[proc_presc_editando, proc_presc_ponte, proc_presc_state,
                 proc_presc_html, proc_presc_editar_btn],
        js=_EDITAR_CRONOLOGIA_JS,
    )
    proc_presc_export_btn.click(
        fn=presc_exportar_html, inputs=[proc_presc_state], outputs=[proc_presc_export_btn]
    )
    proc_perguntar_btn.click(
        fn=proc_responder, inputs=[proc_pergunta, proc_relatorio_state], outputs=[proc_resposta]
    )

    # RJ — análise apenas ao clicar no botão
    rj_analisar_btn.click(
        fn=rj_analisar,
        inputs=[rj_pdf_principal, rj_pdf_relacionados],
        outputs=[rj_log, rj_report, rj_relatorio_state, rj_extracao_state],
        concurrency_limit=20,
        concurrency_id="analises_pdf",
    )
    environment_status_timer.tick(
        fn=environment_status_json,
        outputs=[environment_status_state],
        queue=False,
        show_progress="hidden",
    )
    def _btn_gerando(_texto_normal):
        # Deixa o botão claramente "Gerando..." (desabilitado) enquanto a função roda —
        # sem isso, cliques em botões que chamam a IA (Excel/Checklist) ficavam sem
        # nenhum feedback visual até terminar.
        _ini = lambda: gr.update(value="Gerando...", interactive=False)
        _fim = lambda: gr.update(value=_texto_normal, interactive=True)
        return _ini, _fim

    _word_ini, _word_fim = _btn_gerando("Baixar Word")
    rj_word_btn.click(_word_ini, None, rj_word_btn, queue=False).then(
        fn=rj_gerar_word, inputs=[rj_relatorio_state], outputs=[rj_word_file]
    ).then(_word_fim, None, rj_word_btn, queue=False)

    _excel_ini, _excel_fim = _btn_gerando("Gerar Excel de Credores")
    rj_excel_cred_btn.click(_excel_ini, None, rj_excel_cred_btn, queue=False).then(
        fn=rj_gerar_excel_credores,
        inputs=[rj_relatorio_state, rj_extracao_state],
        outputs=[rj_excel_cred_file, rj_excel_cred_status],
    ).then(_excel_fim, None, rj_excel_cred_btn, queue=False)

    _checklist_ini, _checklist_fim = _btn_gerando("Checklist RJ")
    rj_checklist_btn.click(_checklist_ini, None, rj_checklist_btn, queue=False).then(
        fn=rj_gerar_checklist,
        inputs=[rj_relatorio_state, rj_extracao_state],
        outputs=[rj_checklist_file, rj_checklist_status],
    ).then(_checklist_fim, None, rj_checklist_btn, queue=False)

    def _rj_add_credor(_count):
        _count = min(int(_count) + 1, RJ_MAX_CRED)
        return [_count] + [gr.update(visible=(k < _count)) for k in range(RJ_MAX_CRED)]
    rj_cred_add_btn.click(_rj_add_credor, inputs=[rj_cred_count], outputs=[rj_cred_count] + rj_cred_rows)

    _checklist_cred_ini, _checklist_cred_fim = _btn_gerando("Gerar Checklist de Créditos")
    rj_checklist_cred_btn.click(_checklist_cred_ini, None, rj_checklist_cred_btn, queue=False).then(
        fn=rj_gerar_checklist_creditos,
        inputs=[rj_relatorio_state, rj_extracao_state] + rj_cred_nomes + rj_cred_docs,
        outputs=[rj_checklist_cred_file, rj_checklist_cred_status],
    ).then(_checklist_cred_fim, None, rj_checklist_cred_btn, queue=False)
    rj_perguntar_btn.click(
        fn=rj_responder, inputs=[rj_pergunta, rj_relatorio_state], outputs=[rj_resposta]
    )

    # Matrículas
    def _mat_add_pessoa(contagem):
        contagem = min(int(contagem) + 1, MAT_MAX_PESSOAS)
        return [contagem] + [gr.update(visible=(k < contagem)) for k in range(MAT_MAX_PESSOAS)]

    mat_dev_add_btn.click(_mat_add_pessoa, inputs=[mat_dev_count],
                          outputs=[mat_dev_count] + mat_dev_rows)
    mat_grp_add_btn.click(_mat_add_pessoa, inputs=[mat_grp_count],
                          outputs=[mat_grp_count] + mat_grp_rows)

    # Ordem achatada esperada por _mat_pares_dos_campos: nomes de devedor, docs de
    # devedor, nomes do grupo, docs do grupo — quatro blocos do mesmo tamanho.
    mat_gerar_btn.click(
        fn=mat_gerar_excel,
        inputs=([mat_arquivos, mat_data_ajuizamento]
                + mat_dev_nomes + mat_dev_docs + mat_grp_nomes + mat_grp_docs),
        outputs=[mat_log, mat_status, mat_excel],
        concurrency_limit=2,
    )
    mat_perguntar_btn.click(
        fn=mat_responder, inputs=[mat_pergunta, mat_log], outputs=[mat_resposta]
    )

    # Timeline Societária
    tl_analisar_btn.click(
        fn=timeline_analisar,
        inputs=[tl_arquivos],
        outputs=[tl_status, tl_timeline_html, tl_data_state, tl_acoes_row,
                 tl_tabela_word_file],
        concurrency_limit=2,
    )

    # Edição no próprio painel. Evento único com js= + fn=: o JS liga/desliga o
    # contenteditable e, ao concluir, devolve a árvore serializada, que o Python aplica
    # sobre o JSON e redesenha.
    tl_editar_btn.click(
        fn=timeline_aplicar_html,
        # Só campos simples entram aqui: o JS recebe os inputs e o que ele devolve vira
        # os argumentos do Python, então um gr.State atravessando o JavaScript seria a
        # parte frágil da ponte. O painel já carrega o schema inteiro (inclusive o que
        # não desenha, em data-tl-extra), então o estado anterior não é necessário.
        inputs=[tl_editando, tl_edicao_ponte],
        outputs=[tl_editando, tl_edicao_ponte, tl_data_state, tl_timeline_html, tl_editar_btn],
        js=_EDITAR_TIMELINE_JS,
    )

    tl_exportar_html_btn.click(
        fn=timeline_exportar_html, inputs=[tl_data_state], outputs=[tl_exportar_html_btn]
    )
    tl_gerar_tabela_btn.click(
        fn=lambda data: gr.update(value=timeline_gerar_word(data, None), visible=True),
        inputs=[tl_data_state],
        outputs=[tl_tabela_word_file],
    )

    # Coleta de Informações
    coleta_gerar_dossie_btn.click(
        fn=coleta_gerar_dossie_dispatch,
        inputs=[coleta_excel_in, coleta_excel_coleta_in, coleta_dossie_in],
        outputs=[coleta_log, coleta_status, coleta_dossie_out],
    )
    coleta_gerar_btn.click(
        fn=coleta_gerar,
        inputs=[coleta_excel_in],
        outputs=[coleta_log, coleta_status, coleta_excel_out],
    )


demo.queue(max_size=20)
demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)), max_file_size="2gb", ssr_mode=False)
