---
name: deploy-space
description: Publica o robô de análises jurídicas na ordem segura — testes, app local na versão do Space, CI, Space de teste e só então produção com confirmação da Julia. Use ao terminar uma mudança e antes de qualquer publicação.
---

# Publicar o robô, na ordem

**Regra que não se negocia:** nada vai para produção sem passar pelo Space de teste e sem a
Julia confirmar no chat. Produção é `olivbernardo/analise_matriculas`; teste é
`olivbernardo/analise_matriculas_teste`.

## Degrau 1 — testes na máquina

```bash
.venv_space/Scripts/python.exe -m pytest tests/ -q     # versão do Space (5.29.1) — a que vale
.venv_test/Scripts/python.exe -m pytest tests/ -q      # ambiente novo (gradio 6.x), opcional
```

Se o `.venv_space` não existir:

```bash
uv venv .venv_space --python 3.11
uv pip install --python .venv_space/Scripts/python.exe -r requirements.txt "gradio==5.29.1" pytest
```

Se o `uv` reclamar de "Missing expected target directory for Python minor version link", apontar
direto para o interpretador baixado (`uv python list --only-installed` mostra o caminho).

## Degrau 2 — app rodando local, na versão do Space

Subir com a configuração **"app (versão do Space)"** do `.claude/launch.json` e conferir no
navegador:

- as cinco abas abrem: Processos, Recuperação Judicial, Matrículas, Timeline Societária, Coleta
  de Informações;
- o cabeçalho aparece com o logo e o cartão de status;
- uma análise de verdade inicia e mostra progresso — usar um caso de `../docs teste/processos/`
  ou `../docs teste/rj/`;
- o que a mudança tocou funciona clicando, não só em teste.

## Degrau 3 — CI

`git push` da branch (nunca da `main`). O workflow `testes-e-space-teste.yml` roda `pytest` na
versão de Gradio que o Space usa. **Ler o resultado**; vermelho aqui é informação, não obstáculo
a contornar.

## Degrau 4 — Space de teste

O próprio workflow publica no Space de teste quando os testes passam. Então:

1. Abrir `https://huggingface.co/spaces/olivbernardo/analise_matriculas_teste` e esperar o build.
2. Conferir o que só existe lá: secrets e variáveis, storage persistente
   (`/data/analysis_runtime`), fila e limite de análises, cota real das chaves.
3. **Avisar a Julia**, passar o link e dizer exatamente o que conferir.
4. **Esperar.** Não seguir para produção sem a confirmação dela.

## Produção — só depois da confirmação

Merge na `main`, que dispara o `sync.yml`. Conferir o Space de produção no ar depois.

Se o token de publicação (`HF_TOKEN`) não tiver permissão para o Space de destino, o passo falha
com erro de permissão: **parar e avisar**. Gerar ou trocar token é a Julia quem faz, nas
configurações do Hugging Face.

## Antes de fechar

- `README.md` atualizado se mudou módulo, limite ou variável de ambiente.
- `sdk_version` do README é a fonte única da versão de Gradio: o CI instala essa versão e o
  Space também. Mudar ali muda os dois.
