---
title: Analises Juridicas
emoji: ⚖️
colorFrom: blue
colorTo: indigo
sdk: gradio
sdk_version: "5.29.1"
app_file: app.py
pinned: false
---

## Configuração dos modelos Gemini

Os robôs de Processo e Recuperação Judicial separam os modelos por responsabilidade:

```text
GEMINI_MODEL_EXTRACAO=gemini-3.5-flash-lite
GEMINI_MODEL_RELATORIO=gemini-3.6-flash
GEMINI_MODEL_ESTRUTURADO=gemini-3.5-flash-lite
GEMINI_MODEL_QA=gemini-3.6-flash
```

Esses valores já são os defaults do código. No Hugging Face, podem ser sobrescritos em
`Settings > Variables and secrets > Variables`. As chaves `GEMINI_API_KEY_1`,
`GEMINI_API_KEY_2`, etc. devem continuar cadastradas como secrets.
