---
name: nova-aba
description: Cria um fluxo/aba novo no robô de análises jurídicas seguindo o padrão do repositório, sem deixar para trás os passos que costumam virar fix no commit seguinte. Use quando for adicionar um módulo, uma aba ou um novo botão de análise ao app.py.
---

# Nova aba / novo fluxo de análise

Roteiro fechado. A razão de existir: no histórico deste repositório, quase toda aba nova foi
seguida por um `fix:` de um passo esquecido — `concurrency_limit`, feedback de progresso,
botão que aparecia antes de existir resultado. Percorrer os passos na ordem, sem pular.

A referência viva é a aba **Timeline Societária**: layout em `app.py` (busque
`with gr.Tab("Timeline Societária")`) e wiring logo abaixo do comentário
`# Timeline Societária`. Copiar aquele formato, não inventar outro.

## Antes de escrever código

1. Perguntar, se não estiver claro: o que entra (tipos de arquivo, campos), o que sai (Word,
   Excel, HTML, PNG), e se é análise pesada (vai para a fila) ou operação curta.
2. Ler o `CLAUDE.md` da raiz — em especial os invariantes e a trava das regras jurídicas.
3. Confirmar em qual módulo a lógica mora: aba nova em `app.py` chama função de um módulo
   próprio (`timeline_societaria.py`, `matriculas.py`, …). Interface não carrega regra de
   negócio.

## Layout (dentro do `gr.Blocks`)

4. `with gr.Tab("<Nome>")`, na posição pretendida entre as abas existentes.
5. `gr.Row` com duas colunas: à esquerda o `gr.File` (`file_types`, `file_count`), à direita um
   `gr.Markdown` com **"Como usar:" numerado** — todas as abas têm isso.
6. Linha de ação: `with gr.Row(elem_classes=["analysis-action-row"])` e o botão principal com
   `variant="primary"` e `elem_classes=["analysis-run-btn"]`.
7. `gr.Textbox` de status com `interactive=False`, `elem_classes=["log-area"]` e `placeholder`
   dizendo que o andamento aparece ali.
8. `gr.State` para o resultado estruturado, quando houver edição ou exportação posterior.
9. Ações secundárias (editar, exportar, baixar) dentro de
   `with gr.Row(visible=False) as <pref>_acoes_row` — **elas só existem depois que há
   resultado na tela**; antes disso a linha vazia só ocupa espaço.
10. Caixa de download como `gr.File(interactive=False, visible=False, height=72,
    elem_classes=["compact-file-output"])` — aparece com o arquivo, não antes dele.

## Wiring (na seção de eventos)

11. `.click()` do botão principal **com `concurrency_limit`** (2 nas análises pesadas, como em
    Processos, RJ, Matrículas e Timeline). Esquecer isso já custou um `fix:`.
12. `outputs` inclui status, resultado, `gr.State` e as linhas/caixas que passam a
    `visible=True`.
13. Progresso: a função tem de emitir andamento no status enquanto processa. Botão que fica
    mudo durante a análise é regressão conhecida ("Gerando..." existe por isso).
14. Erro: mostrar **o que a API respondeu**, nunca chutar a causa. E se uma parte obrigatória
    falhar, **bloquear a saída** em vez de entregar resultado parcial como completo.
15. Se houver edição na tela, seguir a ponte da Timeline: evento único com `js=` + `fn=`,
    passando só campos simples (`gr.State` não atravessa o JavaScript com segurança).

## Fechamento — nada disso é opcional

16. CSS novo, se houver, em `assets/design.css` (não em `design.py`, que só monta as strings).
17. Teste em `tests/`, no padrão dos existentes. Se a aba gera documento, incluir asserção de
    rastreabilidade (referência = identificador do tribunal + página absoluta).
18. Rodar a suíte **na versão do Space**:
    `.venv_space/Scripts/python.exe -m pytest tests/ -q`
19. Subir o app local e clicar de verdade, com um caso de `../docs teste/`.
20. Atualizar o `README.md` (a tabela de módulos e, se mudou algum limite, a seção de
    configuração) no mesmo commit.
21. Commit sucinto em pt-BR, sem menção a IA. Depois: branch → CI → Space de teste → **esperar
    a confirmação da Julia** antes de qualquer coisa ir para produção.
