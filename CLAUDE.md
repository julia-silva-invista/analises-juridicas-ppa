# CLAUDE.md — como trabalhar neste repositório

Documento interno de desenvolvimento. Não vai para o Space (é removido no envio).
A descrição do sistema para quem só vai usá-lo está no `README.md`.

## Regra que vale antes de qualquer outra

**Nada vai para produção sem passar pelo Space de teste e sem confirmação explícita da Julia.**
Caminho obrigatório: branch → `pytest` → app local → Space de teste → ela confere e confirma →
produção. Produção é `olivbernardo/analise_matriculas`, publicada pelo push na `main` via
`.github/workflows/sync.yml`. Teste é `olivbernardo/analise_matriculas_teste`.

Commits: pt-BR, **sucintos**, sem `Co-Authored-By` e sem qualquer menção a IA.

## Mapa: sintoma → arquivo

| Sintoma | Onde olhar |
|---|---|
| Citação errada, marcador interno vazando, referência não confirmada | `legal_prompts.py` |
| Fila, limite de análises simultâneas, workers, chunk de PDF | `analysis_runtime.py` |
| Chamada ao Gemini, escolha de modelo, failover de chave, timeout | `utils.py` |
| Interface, aba, botão, wiring de evento | `app.py` |
| Conteúdo/estrutura do relatório de processos | `processos.py`, `report_template_processos.py` |
| Conteúdo/estrutura do relatório de RJ | `rj.py`, `report_template_rj.py` |
| Checklists de RJ em Word | `checklist_rj.py` |
| Dossiês em Word | `dossie_ppa.py` (completo), `dossie_previa.py` (triagem) |
| Matrículas: cadeia dominial, ônus, Excel, destaques | `matriculas.py` |
| Timeline societária: extração, edição, export | `timeline_societaria.py` |
| Cronologia processual (linha do tempo) | `cronologia_prescricao.py` |
| Regime de contagem da prescrição intercorrente | `prescricao_intercorrente.py` |
| Planilha da Coleta de Informações | `coleta.py` (template `x.xlsx`) |
| Reaproveitamento de chunks já extraídos na RJ | `rj_cache.py` |
| CSS, cabeçalho e rodapé | `design.py` + `assets/` |

## Invariantes — não negociar sem pedido explícito

1. **Resultado parcial nunca é apresentado como completo.** Se uma parte obrigatória não pôde ser
   analisada, a geração é interrompida. Nunca "completar" com o que deu.
2. **Erro mostra o que a API respondeu**, em vez de chutar a causa. O log precisa dizer qual
   credencial recusou e por quê.
3. **Referência = identificador real do tribunal + página absoluta do PDF.** Marcador interno de
   parte/chunk jamais aparece no texto final; citação não confirmada é rebaixada para
   "(referência processual não localizada)" — ver `legal_prompts.py:255`.
4. **Fato, indício, hipótese e conclusão são coisas diferentes** e devem continuar distinguíveis na
   saída. Não presumir fraude, grupo econômico ou responsabilidade a partir de vínculo superficial.
5. **Botão de análise novo nasce com `concurrency_limit`** e com feedback de progresso na tela.
   Hoje os cinco botões de análise têm limite (`app.py`, linhas 726, 739, 764, 832, 844).

## Regras jurídicas: trava, com espaço para sugestão

As regras de fidelidade em `legal_prompts.REGRAS_CONSOLIDACAO_PROCESSUAL` são verificadas
**literalmente** por `tests/test_fidelidade_juridica.py` — o teste afirma que trechos exatos
continuam no prompt (`"PÁGINA ABSOLUTA DO PDF"`, `"JAMAIS pode aparecer"`, entre outros).

Portanto: **não alterar essas regras por iniciativa própria.** Sugerir mudança é bem-vindo quando
fizer sentido — mas sempre sinalizando que existe a trava, que os testes quebram junto e que a
decisão é da Julia.

## Números que o código fixa (e o README repete)

| Parâmetro | Valor | Onde |
|---|---|---|
| Análises simultâneas | 4 | `analysis_runtime.py:88` (`MAX_ACTIVE_ANALYSES`) |
| Workers de extração | 6 | `analysis_runtime.py:28` (`TOTAL_EXTRACTION_WORKERS`) |
| Timeout por chamada ao Gemini | 600.000 ms (10 min) | `utils.py:20` (`GEMINI_TIMEOUT_MS`) |
| Modelos por finalidade | 5 variáveis `GEMINI_MODEL_*` | `utils.py:25-32` |
| Chaves Gemini | `GEMINI_API_KEY_1..n`, em ordem, sem pular número | `utils.py:68+` |

Todos são sobrescrevíveis por variável de ambiente no Space. Ao mudar qualquer um, atualizar o
`README.md` no mesmo commit.

## Armadilha do Hugging Face: arquivo binário

O Space **recusa push de arquivo binário** que não esteja em armazenamento LFS/Xet — a mensagem é
`Your push was rejected because it contains binary files`. Vale para o Space de teste e para
produção.

O que passa hoje: `x.xlsx` (declarado como LFS no `.gitattributes`) e os `.docx` de `assets/`, que
são pequenos. Um PNG de 351 KB solto foi rejeitado.

Por isso o logo do cabeçalho é guardado como **texto** em `assets/logo_header.b64`, não como
`.png`. Ao acrescentar imagem, fonte ou qualquer binário, escolher entre: guardar em base64 como
texto, declarar em LFS no `.gitattributes`, ou remover no passo de publicação.

## Comandos

```bash
# Testes (ambiente atual: gradio 6.22.0, Python 3.14.6)
.venv_test/Scripts/python.exe -m pytest tests/ -q

# App local
python app.py            # http://localhost:7861

# Ambiente igual ao do Space (gradio 5.29.1, Python 3.11 baixado pelo uv)
uv venv .venv_space --python 3.11
uv pip install --python .venv_space -r requirements.txt "gradio==5.29.1"
```

Material de teste real (fora do repositório): `../docs teste/processos/` e `../docs teste/rj/`.

## Convenções

- Nomes de função, variável e arquivo em pt-BR, seguindo o que já existe.
- Comentário explica **intenção**, não implementação.
- Sem data de "última atualização" em documento — o histórico é do git.
- `esaj.py` não existe mais (removido em `bf6f7d0`): era código morto desde a saída do modo direto.
