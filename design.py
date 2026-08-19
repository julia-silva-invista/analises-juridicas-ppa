"""Aparencia da interface: CSS, cabecalho e rodape.

O conteudo pesado mora em assets/: o logo do cabecalho (assets/logo_header.png, 351 KB) e a
folha de estilo (assets/design.css, 65 KB). Embutidos no modulo, somavam 549 KB — uma linha
sozinha tinha 478 mil caracteres, o que tornava o arquivo impraticavel de abrir e editar.

A extracao e so de armazenamento: o CSS e o HTML entregues ao Gradio sao identicos aos de
antes, e tests/test_design_snapshot.py guarda a impressao digital das tres saidas. Os
caminhos sao relativos a este modulo, nao ao diretorio de execucao, que no Space e outro.
"""
import base64
from pathlib import Path

_ASSETS = Path(__file__).resolve().parent / 'assets'
_LOGO_BASE64 = base64.b64encode((_ASSETS / 'logo_header.png').read_bytes()).decode('ascii')

CSS = (_ASSETS / 'design.css').read_text(encoding='utf-8')

_HEADER_TEMPLATE = """
<div class="inv-header">
  <div class="inv-header-inner">
    <div class="inv-logo-wrap">
      <img class="inv-logo-img" src="data:image/png;base64,@@LOGO_BASE64@@" alt="Invista" />
    </div>

    <section class="inv-title-block" aria-label="Identificação do produto">
      <div class="inv-eyebrow">Plataforma de automatização de</div>
      <h1>Análises Jurídicas</h1>
      <p class="inv-subtitle">
        Desenvolvida para elaboração de relatórios estruturados de processos diversos, recuperações judiciais e matrículas imobiliárias, para suporte à atividade jurídica.
      </p>
    </section>

    <aside class="inv-status-card" data-status="stable" aria-label="Status operacional">
      <div class="inv-status-top">
        <span class="inv-status-label">Ambiente</span>
        <span class="inv-status-pill"><span class="inv-status-dot"></span> <span class="inv-status-value">Estável</span></span>
      </div>
      <p class="inv-status-text">Selecione o módulo de análise para iniciar a triagem documental e a estruturação do relatório.</p>
    </aside>
  </div>
</div>
"""

HEADER_HTML = _HEADER_TEMPLATE.replace('@@LOGO_BASE64@@', _LOGO_BASE64)

# NOTA: a logica de status/carrossel de abas NAO fica num <script> aqui dentro — HTML
# inserido via innerHTML (que e como gr.HTML() renderiza isso) nunca executa <script>
# tags, em nenhum navegador (confirmado: mesmo o script antigo que só vivia aqui nunca
# rodava de fato). Essa logica agora mora em app.py, passada via `js=` do gr.Blocks —
# o unico mecanismo do Gradio que realmente executa JS no carregamento da pagina.


FOOTER_HTML = '<div class="inv-footer">Invista — Análises Jurídicas · Uso interno restrito</div>'
