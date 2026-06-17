# -*- coding: utf-8 -*-
"""
Análises Jurídicas — Invista
Robô consolidado: Análise de Processos · Recuperação Judicial · Matrículas
"""
import os
import json
import inspect
import requests
from datetime import datetime

import gradio as gr

from design import CSS, HEADER_HTML, FOOTER_HTML
from processos import proc_analisar, proc_gerar_word, proc_responder
from rj import rj_analisar, rj_gerar_word, rj_responder, rj_gerar_excel_credores
from matriculas import mat_gerar_excel, mat_responder
from coleta import coleta_gerar

os.makedirs("resultados", exist_ok=True)
os.makedirs("tmp_pdfs", exist_ok=True)

FEEDBACK_TO       = "juliadeoliveirabernardo@gmail.com"
FEEDBACK_FROM     = "juliadeoliveirabernardo@gmail.com"
SENDGRID_API_KEY  = os.getenv("SENDGRID_API_KEY", "")


def _enviar_email_feedback(nome: str, email: str, mensagem: str, data: str) -> tuple:
    if not SENDGRID_API_KEY:
        return False, "SENDGRID_API_KEY não configurada"
    try:
        corpo = (
            f"<b>Nome:</b> {nome}<br>"
            f"<b>E-mail:</b> {email or '(não informado)'}<br>"
            f"<b>Data:</b> {data}<br><br>"
            f"<b>Mensagem:</b><br>{mensagem.replace(chr(10), '<br>')}"
        )
        payload = {
            "personalizations": [{"to": [{"email": FEEDBACK_TO}]}],
            "from": {"email": FEEDBACK_FROM, "name": "Robô Jurídico"},
            "subject": f"[Feedback] {nome} — {data}",
            "content": [{"type": "text/html", "value": corpo}],
        }
        resp = requests.post(
            "https://api.sendgrid.com/v3/mail/send",
            headers={"Authorization": f"Bearer {SENDGRID_API_KEY}", "Content-Type": "application/json"},
            json=payload,
            timeout=15,
        )
        if resp.status_code in (200, 202):
            return True, ""
        return False, f"Status {resp.status_code}: {resp.text}"
    except Exception as e:
        return False, str(e)


def salvar_feedback(nome: str, email: str, mensagem: str) -> str:
    if not mensagem.strip():
        return "Por favor, escreva sua mensagem antes de enviar."
    nome_exibido = nome.strip() if nome.strip() else "Anônimo"
    email_exibido = email.strip() if email.strip() else ""
    data = datetime.now().strftime("%Y-%m-%d %H:%M")

    enviado, erro = _enviar_email_feedback(nome_exibido, email_exibido, mensagem.strip(), data)

    if enviado:
        return "Feedback enviado com sucesso!"
    elif not SENDGRID_API_KEY:
        return "Feedback recebido (e-mail não configurado)."
    else:
        return f"Não foi possível enviar o e-mail: {erro}"


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
}
"""

with gr.Blocks(
    title="Análises Jurídicas — Invista",
    css=CSS,
    js=_FORCE_LIGHT_JS,
    theme=gr.themes.Default(),
) as demo:
    gr.HTML(HEADER_HTML)

    with gr.Tabs(elem_classes=["tab-nav"]):

        # ── Tab 1: Análise de Processos ──────────────────────────────────────
        with gr.Tab("Processos"):
            with gr.Row(elem_classes=["input-grid-three"]):
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    proc_pdf_principal = gr.File(
                        label="Processo principal",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                        elem_classes=["equal-input-box", "upload-equal-panel"],
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    proc_pdf_relacionados = gr.File(
                        label="Processos relacionados (opcional)",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                        elem_classes=["equal-input-box", "upload-equal-panel"],
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    with gr.Group(elem_classes=["instructions-with-gemini"]):
                        proc_instrucoes = gr.Textbox(
                            label="Instruções adicionais",
                            placeholder="Ex: verificar penhora sobre imóvel; analisar menção a prescrição; checar bem de família...",
                            lines=4,
                            elem_classes=["instructions-field"],
                        )
                        proc_usar_pro = gr.Checkbox(
                            label="Usar Gemini Pro",
                            value=False,
                            elem_classes=["cb-slot", "cb-tip-gemini"],
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

            with gr.Row():
                proc_word_btn = gr.Button("Baixar Word", variant="secondary", elem_classes=["word-download-btn"])
            proc_word_file = gr.File(label="", interactive=False, visible=False, elem_classes=["word-file-output"])

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
            with gr.Row(elem_classes=["input-grid-three"]):
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    rj_pdf_principal = gr.File(
                        label="Processo de RJ",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                        elem_classes=["equal-input-box", "upload-equal-panel"],
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    rj_pdf_relacionados = gr.File(
                        label="Processos relacionados (opcional)",
                        file_types=[".pdf", ".PDF"],
                        file_count="multiple",
                        elem_classes=["equal-input-box", "upload-equal-panel"],
                    )
                with gr.Column(scale=1, elem_classes=["analysis-input-col"]):
                    with gr.Group(elem_classes=["instructions-with-gemini"]):
                        rj_instrucoes = gr.Textbox(
                            label="Instruções adicionais",
                            placeholder="Ex: analisar crédito do Bradesco; verificar prescrição; identificar créditos extraconcursais...",
                            lines=4,
                            elem_classes=["instructions-field"],
                        )
                        rj_usar_pro = gr.Checkbox(
                            label="Usar Gemini Pro",
                            value=False,
                            elem_classes=["cb-slot", "cb-tip-gemini"],
                        )
                        rj_versao_resumida = gr.Checkbox(
                            label="Versão resumida",
                            value=False,
                            elem_classes=["cb-slot", "cb-tip-resumida"],
                        )

            with gr.Row(elem_classes=["analysis-action-row"]):
                rj_analisar_btn = gr.Button("Analisar recuperação judicial", variant="primary", elem_classes=["analysis-run-btn"])

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
                rj_word_btn          = gr.Button("Baixar Word",            variant="secondary", elem_classes=["word-download-btn"])
                rj_excel_cred_btn    = gr.Button("Gerar Excel de Credores", variant="secondary", elem_classes=["word-download-btn"])
            rj_word_file      = gr.File(label="",                  interactive=False, visible=False, elem_classes=["word-file-output"])
            rj_excel_cred_file   = gr.File(label="Excel de Credores", interactive=False, visible=False, elem_classes=["word-file-output"])
            rj_excel_cred_status = gr.Textbox(label="", interactive=False, lines=1, show_label=False)

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

            with gr.Accordion("Dados da execução — alertas de coloração (opcional)", open=False):
                with gr.Row():
                    mat_data_ajuizamento = gr.Textbox(
                        label="Data do ajuizamento",
                        placeholder="DD/MM/AAAA",
                        lines=1,
                        scale=1,
                    )
                with gr.Row():
                    mat_devedores = gr.Textbox(
                        label="Devedores (um por linha: CPF/CNPJ — Nome)",
                        placeholder="123.456.789-00 — João da Silva\n12.345.678/0001-90 — Empresa XYZ Ltda",
                        lines=4,
                        scale=1,
                    )
                    mat_relacionados = gr.Textbox(
                        label="Pessoas do grupo econômico (um por linha: CPF/CNPJ — Nome)",
                        placeholder="111.222.333-44 — Maria Oliveira\n55.666.777/0001-88 — Holdings XYZ S/A",
                        lines=4,
                        scale=1,
                    )
                gr.Markdown(
                    "🟡 **Amarelo claro** — transmissão envolvendo devedor ou pessoa do grupo  \n"
                    "🔴 **Vermelho claro** — transmissão após data de ajuizamento"
                )

            mat_gerar_btn = gr.Button("Gerar Excel", variant="primary", size="lg")

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
                    mat_excel = gr.File(label="Excel gerado", interactive=False)

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

        # ── Tab 4: Coleta de Informações ─────────────────────────────────────
        with gr.Tab("Coleta de Informações"):
            with gr.Row():
                with gr.Column(scale=2):
                    coleta_excel_in = gr.File(
                        label="Excel(s) da Predictus",
                        file_types=[".xlsx", ".xls"],
                        file_count="multiple",
                    )
                with gr.Column(scale=1):
                    gr.Markdown(
                        "**Como usar:**\n\n"
                        "1. Exporte o dossiê jurídico da Predictus em Excel\n"
                        "2. Faça upload de um ou mais arquivos\n"
                        "3. Clique em **Gerar Planilha**\n"
                        "4. Baixe o arquivo preenchido\n\n"
                        "_Suporta múltiplos devedores simultaneamente._\n\n"
                        "_Os processos são distribuídos automaticamente entre as abas Trabalhista e Fiscal & Cível._"
                    )

            coleta_gerar_btn = gr.Button("Gerar Planilha", variant="primary", size="lg")

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
                    coleta_excel_out = gr.File(label="Planilha preenchida", interactive=False)

        # ── Tab 5: Sugestões e Feedbacks ─────────────────────────────────────
        with gr.Tab("Feedbacks"):
            gr.HTML("""
            <div class="feedback-intro">
              <h2>Sugestões e Feedbacks</h2>
              <p>Colabore com a evolução da plataforma e nos ajude a identificar pontos de melhoria, corrigir erros e desenvolver novas funcionalidades.</p>
              <p>Use este espaço para:</p>
              <ul>
                <li>Reportar erros ou inconsistências nas análises</li>
                <li>Sugerir novas funcionalidades ou tipos de análise</li>
                <li>Tirar dúvidas sobre o funcionamento</li>
                <li>Qualquer outro feedback que queira compartilhar</li>
              </ul>
            </div>
            """)

            with gr.Row():
                with gr.Column(scale=1):
                    fb_nome = gr.Textbox(
                        label="Nome (opcional)",
                        placeholder="Seu nome...",
                        lines=1,
                    )
                with gr.Column(scale=1):
                    fb_email = gr.Textbox(
                        label="E-mail (opcional)",
                        placeholder="seu.email@exemplo.com",
                        lines=1,
                    )
            fb_mensagem = gr.Textbox(
                label="Mensagem",
                placeholder="Descreva o erro, sugestão ou dúvida...",
                lines=8,
            )
            fb_enviar_btn = gr.Button("Enviar", variant="primary")
            fb_status = gr.Textbox(label="", interactive=False, lines=1, show_label=False)

    gr.HTML(FOOTER_HTML)

    # ── Eventos ──────────────────────────────────────────────────────────────

    # Processos — análise apenas ao clicar no botão
    proc_analisar_btn.click(
        fn=proc_analisar,
        inputs=[proc_pdf_principal, proc_pdf_relacionados, proc_instrucoes, proc_usar_pro, proc_versao_resumida],
        outputs=[proc_log, proc_report, proc_relatorio_state],
        concurrency_limit=2,
    )
    proc_word_btn.click(fn=proc_gerar_word, inputs=[proc_relatorio_state], outputs=[proc_word_file])
    proc_perguntar_btn.click(
        fn=proc_responder, inputs=[proc_pergunta, proc_relatorio_state], outputs=[proc_resposta]
    )

    # RJ — análise apenas ao clicar no botão
    rj_analisar_btn.click(
        fn=rj_analisar,
        inputs=[rj_pdf_principal, rj_pdf_relacionados, rj_instrucoes, rj_usar_pro, rj_versao_resumida],
        outputs=[rj_log, rj_report, rj_relatorio_state, rj_extracao_state],
        concurrency_limit=2,
    )
    rj_word_btn.click(fn=rj_gerar_word, inputs=[rj_relatorio_state], outputs=[rj_word_file])
    rj_excel_cred_btn.click(
        fn=rj_gerar_excel_credores,
        inputs=[rj_relatorio_state, rj_extracao_state],
        outputs=[rj_excel_cred_file, rj_excel_cred_status],
    )
    rj_perguntar_btn.click(
        fn=rj_responder, inputs=[rj_pergunta, rj_relatorio_state], outputs=[rj_resposta]
    )

    # Matrículas
    mat_gerar_btn.click(
        fn=mat_gerar_excel,
        inputs=[mat_arquivos, mat_data_ajuizamento, mat_devedores, mat_relacionados],
        outputs=[mat_log, mat_status, mat_excel],
        concurrency_limit=2,
    )
    mat_perguntar_btn.click(
        fn=mat_responder, inputs=[mat_pergunta, mat_log], outputs=[mat_resposta]
    )

    # Coleta de Informações
    coleta_gerar_btn.click(
        fn=coleta_gerar,
        inputs=[coleta_excel_in],
        outputs=[coleta_log, coleta_status, coleta_excel_out],
    )

    # Feedbacks
    fb_enviar_btn.click(fn=salvar_feedback, inputs=[fb_nome, fb_email, fb_mensagem], outputs=[fb_status])


demo.queue(max_size=20)
demo.launch(server_name="0.0.0.0", server_port=int(os.getenv("PORT", 7860)), max_file_size="2gb", ssr_mode=False)
