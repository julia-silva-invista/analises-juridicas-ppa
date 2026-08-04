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

Os robôs de Processo, Recuperação Judicial e Matrículas separam os modelos por responsabilidade:

```text
GEMINI_MODEL_EXTRACAO=gemini-3.5-flash-lite
GEMINI_MODEL_RELATORIO=gemini-3.6-flash
GEMINI_MODEL_ESTRUTURADO=gemini-3.5-flash-lite
GEMINI_MODEL_QA=gemini-3.6-flash
```

Esses valores já são os defaults do código. No Hugging Face, podem ser sobrescritos em
`Settings > Variables and secrets > Variables`. As chaves `GEMINI_API_KEY_1`,
`GEMINI_API_KEY_2`, etc. devem continuar cadastradas como secrets.

Em Matrículas, `GEMINI_MODEL_EXTRACAO` faz a leitura factual integral do PDF e
`GEMINI_MODEL_RELATORIO` executa a consolidação jurídica estruturada (cadeia dominial,
proporções, situação da matrícula e classificação dos ônus).

## Limites de processamento

Os valores abaixo também já são os defaults e só precisam ser cadastrados como Variables no
Hugging Face se for necessário alterá-los:

```text
MAX_ACTIVE_ANALYSES=4
TOTAL_EXTRACTION_WORKERS=6
PDF_PREPARATION_CONCURRENCY=2
PDF_CHUNK_MAX_MB=45
ANALYSIS_CACHE_RETENTION_DAYS=7
```

Os seis workers são compartilhados dinamicamente: uma análise pode usar 6; duas usam até 3
cada; três usam até 2 cada; quatro usam 1 cada. A quinta análise permanece na fila visível.
Os PDFs são divididos em faixas de até 400 páginas e 45 MB. Antes da medição, o robô
remove recursos globais não utilizados, preservando texto, vetores e imagens. Somente uma
página isolada que permaneça acima de 45 MB é rasterizada, com redução progressiva de
resolução. O cache e o último estágio ficam em `/data/analysis_runtime` quando o Space possui
storage persistente, ou em `resultados/analysis_runtime` no ambiente local.
