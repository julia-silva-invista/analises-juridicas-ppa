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
    COLUNAS_MARCOS,
    aplicar_marcos as presc_aplicar_marcos,
    exportar_html as presc_exportar_html,
)
from prescricao_intercorrente import TIPOS_DE_MARCO
from rj import rj_analisar, rj_gerar_word, rj_responder, rj_gerar_excel_credores, rj_gerar_checklist, rj_gerar_checklist_creditos
from matriculas import mat_gerar_excel, mat_responder
from coleta import coleta_gerar, coleta_gerar_dossie_dispatch
from analysis_runtime import environment_status_json
from timeline_societaria import (
    timeline_analisar,
    timeline_salvar_imagem,
    timeline_gerar_word,
    timeline_exportar_html,
    aplicar_cabecalho as timeline_aplicar_cabecalho,
    aplicar_edicao as timeline_aplicar_edicao,
    adicionar_evento as timeline_adicionar_evento,
    ordenar_eventos as timeline_ordenar_eventos,
    remover_evento as timeline_remover_evento,
    selecionar_evento as timeline_selecionar_evento,
    COLUNAS_ADMINISTRADORES,
    COLUNAS_CESSOES,
    COLUNAS_FILIAIS,
    COLUNAS_IMOVEIS,
    COLUNAS_SOCIOS,
    MOVIMENTOS_IMOVEL,
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

# Exportação de imagem da Timeline Societária: rasteriza o PRÓPRIO HTML da timeline
# (o mesmo que está na tela) num canvas, via <foreignObject> de SVG — sem biblioteca
# externa e sem redesenhar o layout no servidor.
#
# Roda como pré-processamento do clique (js + fn no MESMO evento): recebe o valor atual
# do campo-ponte e devolve `[data URL]`, que o Gradio entrega como argumento de
# `timeline_salvar_imagem`. Evite `fn=None` + `.then(...)`: encadear um evento de backend
# depois de um evento só-JS depende de comportamento que muda entre a versão de
# desenvolvimento (Gradio 6) e a 5.29.1 do Space. O evento único não depende disso.
# Se qualquer etapa falhar, devolve [""] e o Python explica o que houve.
_CAPTURAR_TIMELINE_JS = """
async (_captura) => {
    const vazio = [""];
    try {
        const shell = document.querySelector("#timeline-export-area");
        if (!shell) return vazio;

        const estilo = shell.querySelector("#tl2-export-style");
        const css = estilo ? estilo.textContent : "";
        const overrides = `
            .tl2-export-clone { width: max-content !important; box-shadow: none !important;
                                margin: 0 !important; background: #fff !important; }
            .tl2-export-clone .tl2-scroll { overflow: visible !important; }
            .tl2-export-clone, .tl2-export-clone * {
                font-family: Arial, Helvetica, sans-serif !important; }
            .tl2-export-clone .tl2-btn,
            .tl2-export-clone .tl2-modal,
            .tl2-export-clone .tl2-toggle { display: none !important; }
        `;

        const clone = shell.cloneNode(true);
        clone.removeAttribute("id");
        clone.classList.add("tl2-export-clone");
        clone.querySelectorAll("#tl2-export-style, .tl2-btn, .tl2-modal, .tl2-toggle")
             .forEach((n) => n.remove());

        // Mede fora da tela, com o eixo horizontal inteiro visível (sem a barra de rolagem).
        const holder = document.createElement("div");
        holder.style.cssText = "position:fixed;left:-100000px;top:0;width:max-content;background:#fff;";
        const estiloMedicao = document.createElement("style");
        estiloMedicao.textContent = css + overrides;
        holder.appendChild(estiloMedicao);
        holder.appendChild(clone);
        document.body.appendChild(holder);
        const rect = clone.getBoundingClientRect();
        const largura = Math.ceil(rect.width);
        const altura = Math.ceil(rect.height) + 8;
        document.body.removeChild(holder);
        if (!largura || !altura) return vazio;

        const wrapper = document.createElement("div");
        wrapper.style.cssText = "width:" + largura + "px;background:#ffffff;";
        const estiloSvg = document.createElement("style");
        estiloSvg.textContent = css + overrides;
        wrapper.appendChild(estiloSvg);
        wrapper.appendChild(clone);
        const xhtml = new XMLSerializer().serializeToString(wrapper);

        const escala = Math.max(1, Math.min(2, 3200 / largura));
        const svg =
            '<svg xmlns="http://www.w3.org/2000/svg" width="' + Math.round(largura * escala) +
            '" height="' + Math.round(altura * escala) + '" viewBox="0 0 ' + largura + " " + altura + '">' +
            '<foreignObject x="0" y="0" width="' + largura + '" height="' + altura + '">' +
            xhtml + "</foreignObject></svg>";

        const imagem = new Image();
        imagem.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
        await new Promise((resolve, reject) => {
            imagem.onload = resolve;
            imagem.onerror = () => reject(new Error("svg"));
            window.setTimeout(() => reject(new Error("timeout")), 20000);
        });

        const canvas = document.createElement("canvas");
        canvas.width = Math.round(largura * escala);
        canvas.height = Math.round(altura * escala);
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(imagem, 0, 0, canvas.width, canvas.height);
        return [canvas.toDataURL("image/png")];
    } catch (e) {
        return vazio;
    }
}
"""

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
            with gr.Accordion("Editar cronologia", open=False, elem_classes=["mat-accordion"]):
                proc_presc_titulo = gr.Textbox(
                    label="Título / lastro (define o prazo aplicado)",
                    placeholder="Ex: Cédula de Crédito Bancário (CCB) nº 123",
                )
                proc_presc_marcos = gr.Dataframe(
                    headers=COLUNAS_MARCOS, col_count=(len(COLUNAS_MARCOS), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label=f"Marcos (tipos: {' · '.join(TIPOS_DE_MARCO)})",
                    elem_classes=["tl-editor-table"],
                )
                proc_presc_export_btn = gr.DownloadButton("Exportar cronologia (HTML)", variant="secondary")

            gr.HTML('<hr class="inv-divider">')
            with gr.Column(elem_classes=["qa-section"]):
                with gr.Row():
                    with gr.Column(scale=4):
                        proc_pergunta = gr.Textbox(
                            label="Perguntas sobre a análise:",
                            lines=3,
                            placeholder="Exemplos: Quando foi realizado o primeiro pedido de penhora? Há risco de prescrição intercorrente? Há risco de sucumbência? Qual o valor atualizado da dívida?"
                        )
                    with gr.Column(scale=1):
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
                with gr.Row():
                    with gr.Column(scale=4):
                        rj_pergunta = gr.Textbox(
                            label="Perguntas sobre a análise:",
                            lines=3,
                            placeholder="Exemplos: Qual é o maior credor e qual percentual detém dos créditos? Qual o status atual do stay period? Há risco de prescrição intercorrente nas execuções?"
                        )
                    with gr.Column(scale=1):
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
                    "🔴 **Vermelho claro** — transmissão após data de ajuizamento  \n"
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
                with gr.Row():
                    with gr.Column(scale=4):
                        mat_pergunta = gr.Textbox(
                            label="Perguntas sobre as matrículas:",
                            lines=3,
                            placeholder="Exemplos: Indique as matrículas que pertencem a determinado devedor. Quais têm penhora vigente? Qual o valor total dos ônus?"
                        )
                    with gr.Column(scale=1):
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
                        "3. Edite os eventos se necessário\n"
                        "4. Exporte como HTML interativo, imagem (A4) ou tabela em Word\n\n"
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
            tl_evento_idx = gr.State(0)

            # Edição por ato: campos rotulados, uma tabelinha por lista do JSON.
            # A grade de 12 colunas anterior exigia decorar uma sintaxe ("Nome | 50%;
            # Outro | 40%", "Cedente > Cessionário | % | valor") e perdia campos no
            # caminho de volta.
            with gr.Accordion("Editar timeline", open=False, elem_classes=["mat-accordion"]):
                with gr.Row():
                    tl_empresa = gr.Textbox(label="Empresa", scale=3)
                    tl_cnpj = gr.Textbox(label="CNPJ", scale=2)

                tl_evento_sel = gr.Dropdown(
                    label="Ato em edição", choices=[], value=None,
                    interactive=True, filterable=False,
                )
                with gr.Row():
                    tl_add_evento_btn = gr.Button("+ Adicionar ato", variant="secondary")
                    tl_del_evento_btn = gr.Button("Remover ato", variant="secondary")
                    tl_ordenar_btn = gr.Button("Ordenar por data", variant="secondary")

                with gr.Row():
                    tl_f_data = gr.Textbox(label="Data", placeholder="DD/MM/AAAA", scale=1)
                    tl_f_ato = gr.Textbox(label="Ato / ACS", placeholder="Ex: 2ª Alteração", scale=2)
                    tl_f_arquivamento = gr.Textbox(label="Nº de arquivamento", scale=1)
                tl_f_detalhamento = gr.Textbox(label="Detalhamento", lines=3)
                with gr.Row():
                    tl_f_cap_ant = gr.Textbox(label="Capital social anterior")
                    tl_f_cap_apos = gr.Textbox(label="Capital social após o ato")
                with gr.Row():
                    tl_f_sede = gr.Textbox(label="Sede após o ato")
                    tl_f_objeto = gr.Textbox(label="Objeto social após o ato")
                tl_f_fonte = gr.Textbox(label="Fonte")

                tl_t_socios = gr.Dataframe(
                    headers=COLUNAS_SOCIOS, col_count=(len(COLUNAS_SOCIOS), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label="Sócios após o ato", elem_classes=["tl-editor-table"],
                )
                tl_t_admin = gr.Dataframe(
                    headers=COLUNAS_ADMINISTRADORES, col_count=(len(COLUNAS_ADMINISTRADORES), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label="Administração após o ato", elem_classes=["tl-editor-table"],
                )
                tl_t_cessoes = gr.Dataframe(
                    headers=COLUNAS_CESSOES, col_count=(len(COLUNAS_CESSOES), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label="Cessões de quotas", elem_classes=["tl-editor-table"],
                )
                tl_t_imoveis = gr.Dataframe(
                    headers=COLUNAS_IMOVEIS, col_count=(len(COLUNAS_IMOVEIS), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label=f"Imóveis (movimento: {' · '.join(MOVIMENTOS_IMOVEL)})",
                    elem_classes=["tl-editor-table"],
                )
                tl_t_filiais = gr.Dataframe(
                    headers=COLUNAS_FILIAIS, col_count=(len(COLUNAS_FILIAIS), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label="Filiais existentes após o ato", elem_classes=["tl-editor-table"],
                )
                tl_t_filiais_novas = gr.Dataframe(
                    headers=COLUNAS_FILIAIS, col_count=(len(COLUNAS_FILIAIS), "fixed"),
                    row_count=(1, "dynamic"), interactive=True, wrap=True,
                    label="Filiais abertas NESTE ato (é o que gera o card 'Filial aberta')",
                    elem_classes=["tl-editor-table"],
                )

            with gr.Row():
                tl_exportar_html_btn = gr.DownloadButton("Exportar HTML", variant="secondary")
                tl_gerar_tabela_btn = gr.Button("Gerar tabela (Word)", variant="secondary")
                tl_exportar_img_btn = gr.Button("Exportar imagem", variant="secondary")

            with gr.Row():
                tl_tabela_word_file = gr.File(
                    label="Tabela em Word", interactive=False, visible=True, height=72,
                    elem_classes=["compact-file-output"],
                )
                tl_imagem_file = gr.File(
                    label="Imagem da timeline", interactive=False, visible=True, height=72,
                    elem_classes=["compact-file-output"],
                )

            # Ponte da captura feita no navegador (data URL PNG) até o Python.
            tl_imagem_captura = gr.Textbox(visible=False)

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
        outputs=[proc_presc_state, proc_presc_html, proc_presc_marcos, proc_presc_titulo],
        concurrency_limit=2,
    )
    # Editar título ou marcos refaz o enquadramento na hora. `.input` e não `.change`:
    # o preenchimento programático do dataframe reescreveria o estado sozinho.
    for _campo in (proc_presc_titulo, proc_presc_marcos):
        _campo.input(
            fn=presc_aplicar_marcos,
            inputs=[proc_presc_state, proc_presc_titulo, proc_presc_marcos],
            outputs=[proc_presc_state, proc_presc_html],
            show_progress="hidden",
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
    # Campos do editor, na MESMA ordem que timeline_societaria espera
    # (CAMPOS_TEXTO_DO_EVENTO seguido de LISTAS_DO_EVENTO).
    tl_campos_evento = [
        tl_f_data, tl_f_ato, tl_f_arquivamento, tl_f_detalhamento,
        tl_f_cap_ant, tl_f_cap_apos, tl_f_sede, tl_f_objeto, tl_f_fonte,
        tl_t_socios, tl_t_admin, tl_t_cessoes, tl_t_imoveis,
        tl_t_filiais, tl_t_filiais_novas,
    ]

    tl_analisar_btn.click(
        fn=timeline_analisar,
        inputs=[tl_arquivos],
        outputs=[tl_status, tl_timeline_html, tl_evento_sel, tl_data_state],
        concurrency_limit=2,
    ).then(
        fn=timeline_selecionar_evento,
        inputs=[tl_data_state, tl_evento_idx],
        outputs=[tl_evento_idx] + tl_campos_evento,
    ).then(
        fn=lambda data: (data.get("empresa", ""), data.get("cnpj", "")),
        inputs=[tl_data_state],
        outputs=[tl_empresa, tl_cnpj],
    )

    # Preview ao vivo: cada campo grava no ato selecionado e redesenha o HTML na hora,
    # em vez de só ao fechar um modo de edição.
    #
    # `.input` e não `.change`: change dispara também quando o Gradio preenche o campo
    # por conta própria. Ao trocar de ato, os 15 campos seriam reescritos — e, se algum
    # deles rodasse antes de tl_evento_idx chegar ao valor novo, gravaria os dados do
    # ato recém-aberto por cima do anterior.
    for _campo in tl_campos_evento:
        _campo.input(
            fn=timeline_aplicar_edicao,
            inputs=[tl_data_state, tl_evento_idx] + tl_campos_evento,
            outputs=[tl_data_state, tl_timeline_html, tl_evento_sel],
            show_progress="hidden",
        )

    for _campo in (tl_empresa, tl_cnpj):
        _campo.input(
            fn=timeline_aplicar_cabecalho,
            inputs=[tl_data_state, tl_empresa, tl_cnpj],
            outputs=[tl_data_state, tl_timeline_html],
            show_progress="hidden",
        )

    tl_evento_sel.select(
        fn=lambda data, evt: timeline_selecionar_evento(data, evt.index),
        inputs=[tl_data_state, tl_evento_sel],
        outputs=[tl_evento_idx] + tl_campos_evento,
    )

    for _botao, _fn in ((tl_add_evento_btn, timeline_adicionar_evento),
                        (tl_del_evento_btn, timeline_remover_evento),
                        (tl_ordenar_btn, timeline_ordenar_eventos)):
        _entradas = [tl_data_state] if _fn is timeline_adicionar_evento else [tl_data_state, tl_evento_idx]
        _botao.click(
            fn=_fn,
            inputs=_entradas,
            outputs=[tl_data_state, tl_timeline_html, tl_evento_sel, tl_evento_idx] + tl_campos_evento,
        )
    # Um único evento: o JS roda antes e o que ele devolve vira o argumento do Python.
    tl_exportar_img_btn.click(
        fn=timeline_salvar_imagem,
        inputs=[tl_imagem_captura],
        outputs=[tl_imagem_file],
        js=_CAPTURAR_TIMELINE_JS,
    )
    tl_exportar_html_btn.click(
        fn=timeline_exportar_html, inputs=[tl_data_state], outputs=[tl_exportar_html_btn]
    )
    tl_gerar_tabela_btn.click(
        fn=lambda data: timeline_gerar_word(data, None),
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
