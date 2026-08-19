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

# Plataforma de Análises Jurídicas — PPA Invista

Conjunto de robôs que leem processos, recuperações judiciais, matrículas imobiliárias e atos
societários e devolvem análise jurídica estruturada, com referência rastreável ao documento de
origem. Serve à prospecção e precificação de ativos: o resultado alimenta pareceres, dossiês e
checklists usados na decisão de aquisição de crédito.

Feito em Python, interface em Gradio, modelos Gemini pela biblioteca `google-genai`. Roda como Space
do Hugging Face — o acesso é pelo navegador, e o envio do código ao GitHub publica a versão em
produção.

## Módulos

| Aba | O que faz | Saídas |
|---|---|---|
| **Processos** | Analisa execução e processos relacionados (incidentes, recursos, correlatos) | Relatório em Word, Dossiê Prévia, Dossiê Desalinhado, Cronologia Processual (HTML/PNG), perguntas sobre o relatório |
| **Recuperação Judicial** | Analisa a RJ e os créditos ligados a ela, cruzando com os processos correlatos | Relatório em Word, Checklist RJ, Checklist de Créditos por credor, Excel de credores (em testes) |
| **Matrículas** | Consolida matrículas em planilha: cadeia dominial, ônus, garantias | Excel com destaques de risco |
| **Timeline Societária** | Reconstrói em ordem cronológica a situação da empresa após cada ato societário | HTML interativo editável, PNG, tabela em Word |
| **Coleta de Informações** | Preenche a planilha-modelo `x.xlsx` a partir de extrações da Predictus | Planilha preenchida, Dossiê atualizado |

### Processos
Recebe o processo principal e, opcionalmente, incidentes, recursos e correlatos. Aceita **Instruções
Adicionais**, que têm precedência sobre o template padrão, alcançam os dossiês e podem gerar quadro
próprio no documento final. Levanta qualificação das partes, advogados, distribuição, valor da causa,
lastro, garantias, assinaturas, índices, citações, embargos, recursos e principais andamentos, tudo
consolidado em template jurídico próprio para reduzir variação entre análises equivalentes.

A **Cronologia Processual** monta a linha do tempo dos marcos que importam para prescrição
intercorrente, editável na própria tela (incluir e excluir marcos; ao concluir, o servidor reordena
por data e recalcula) e exportável em HTML autocontido ou PNG.

Divisão de trabalho deliberada: **o modelo extrai apenas fatos datados**; o enquadramento — qual
regime se aplica, o que suspende ou zera a contagem — é calculado depois em Python determinístico,
sobre base de regras curada (`prescricao_intercorrente.py`), considerando CPC/1973, CPC/2015,
Lei 14.195/2021, CC/1916 e CC/2002. O reconhecimento da prescrição não é conclusão autônoma do
modelo.

Dossiês e cronologia usam o texto já extraído como fonte — não reabrem o PDF.

### Recuperação Judicial
Analisa o processo principal da RJ e documentos de créditos específicos (execuções, habilitações,
divergências, impugnações), cruzando automaticamente o que veio dos autos da RJ com o que veio dos
processos correlatos para identificar o credor a mapear. Também aceita os dados do credor informados
diretamente, para o caso em que o crédito não foi discutido em outro processo.

Levanta: recuperandas, administrador judicial, pedido e deferimento, perícia prévia, consolidação
substancial, RMA, QGC, PRJ, condições de pagamento por classe, AGC, stay period e essencialidade de
bens. Por credor, reconstrói a evolução do crédito entre editais, divergência administrativa, posição
do administrador judicial, impugnação, decisões, recursos, garantias e o confronto entre a
classificação das recuperandas e a sustentada pelo credor.

O relatório consolidado pode ser reapresentado junto de novos processos relacionados, reaproveitando
a análise já feita em vez de reextrair os autos inteiros.

### Matrículas
Consolida cada matrícula em uma linha estruturada no Excel: número e situação, cartório, comarca,
proprietário atual, descrição, fração ideal, cadeia de transmissões, ônus vigentes e cancelados,
garantias, grau de confiança e valor total dos ônus, ao lado dos campos de avaliação (VM e VP)
reservados ao analista. Para ônus financeiros, quando disponíveis: credor, processo ou instrumento,
valor, moeda, datas de constituição e vencimento, cancelamento e o código de registro ou averbação
(R./AV.) de origem, que permite localizar o gravame na matrícula.

O fluxo tem duas etapas: o PyMuPDF inspeciona a camada textual e direciona ao modelo de extração
textual ou ao de OCR; depois, uma segunda chamada consolida a cadeia dominial e a situação dos ônus.
As respostas de consolidação seguem schemas Pydantic, revalidados no retorno.

O que **não** fica a cargo do modelo: confronto de nomes e CPF/CNPJ das partes das transmissões, e a
apuração do valor dos ônus (atualização a 1% ao mês, juros simples, do ato até a data da análise) e
do saldo final. As matrículas são processadas em paralelo; `pandas` organiza a base e `openpyxl` gera
e formata o arquivo, destacando:

- **amarelo** — transmissões que envolvam o devedor ou pessoas do grupo, quando indicados;
- **vermelho** — especificamente aquelas em que o próprio devedor figura como transmitente em data
  igual ou posterior ao ajuizamento da execução, critério do art. 792 do CPC (sem prejuízo da
  verificação dos demais requisitos).

### Timeline Societária
Todos os atos são submetidos em **uma única chamada multimodal**, para o modelo ler a sequência como
conjunto e não documento a documento. Identifica mudanças no quadro societário e administrativo,
participações e cessões de quotas, capital social, sede, objeto social, nome empresarial, filiais e
movimentações de imóveis, registrando data, natureza, número de arquivamento e páginas de origem de
cada ato.

O retorno vem em JSON; rotinas em Python comparam cada evento com o estado anterior para identificar
o que mudou de fato. A timeline é exibida em página interativa (HTML + CSS, edição direta via
JavaScript); ao concluir a edição, o conteúdo volta ao servidor, que reordena por data e recalcula
antes de redesenhar. Exporta HTML interativo, PNG ou tabela em Word.

## Arquitetura

O processamento é dividido em camadas: o frontend recebe arquivos e parâmetros; o backend prepara e
segmenta os documentos; os modelos extraem e consolidam; rotinas em Python verificam consistência,
aplicam classificações e cálculos objetivos e organizam as entregas em Word, Excel, HTML ou PNG.

### Controle de carga
Processos e RJ comportam **quatro análises ativas simultâneas**; as demais entram em fila visível,
com indicação de posição. Os módulos mais curtos rodam fora dessa fila.

São **seis workers de extração compartilhados**, distribuídos conforme a demanda: uma análise pode
usar 6; duas usam até 3 cada; três usam até 2 cada; quatro usam 1 cada. A quinta análise permanece na
fila. Quem coordena é o `AnalysisManager` (`analysis_runtime.py`); o paralelismo usa
`concurrent.futures.ThreadPoolExecutor`.

O cabeçalho da interface mostra em tempo real o estado do ambiente, calculado pela soma das análises
ativas e em fila: até 2 é "Estável", 3 é "Operacional", 4 ou mais é "Pico de demanda".

A fila é usada apenas por Processos e Recuperação Judicial (`analysis_runtime.ANALYSIS_MANAGER`);
Matrículas, Timeline Societária e Coleta rodam fora dela.

### Segmentação de documentos extensos
PDFs de processos e RJ podem ter milhares de páginas, muitas escaneadas. Em vez de uma chamada única,
os documentos são divididos em partes (*chunks*) de até **400 páginas** e **45 MB**. Antes da
medição, o robô remove recursos globais não utilizados, preservando texto, vetores e imagens.

Dois limites diferentes atuam em momentos distintos:

- **Tamanho** — verificado localmente, antes do envio. Parte que excede é subdividida. Página isolada
  que permaneça acima do limite é rasterizada com resolução decrescente (180, 144 e 108 DPI) até
  caber.
- **Contexto do modelo** — só se revela na recusa da API. O sistema identifica a recusa pela mensagem
  retornada, parte aquele trecho ao meio e o reprocessa, com as páginas reindexadas para a numeração
  absoluta do PDF original.

### Página pesquisável vs. escaneada
O PyMuPDF inspeciona a camada textual. Uma página só é considerada pesquisável se tiver **ao menos 50
caracteres, 30 palavras e 40% de caracteres alfabéticos** (`utils.py`). O critério evita que
elementos isolados — o número do processo no rodapé, por exemplo — façam uma página escaneada passar
por pesquisável. Páginas com texto confiável vão ao modelo de extração leve, otimizado para volume;
trechos com páginas escaneadas vão ao modelo multimodal de OCR.

### Rastreabilidade: instrução **e** verificação em código
As informações do relatório são vinculadas ao identificador real usado pelo tribunal e à **página
absoluta do PDF original**, sem exibir divisões técnicas internas do processamento.

Isso não depende só da instrução dada ao modelo. Uma camada determinística em Python, aplicada sobre
a resposta pronta (`legal_prompts.py`), elimina marcador interno de processamento eventualmente
vazado e rebaixa citação de página não confirmada para a advertência "(referência processual não
localizada)".

E se alguma parte obrigatória do documento não puder ser analisada, **a geração do relatório é
interrompida** — resultado parcial não é apresentado como se todo o processo tivesse sido lido.

## Configuração

### Modelos Gemini
Separados por responsabilidade, para usar modelo leve em tarefa de volume e modelo robusto em leitura
visual e consolidação jurídica. Os valores abaixo já são os defaults do código e podem ser
sobrescritos no Space em `Settings > Variables and secrets > Variables`, sem alterar o código:

```text
GEMINI_MODEL_EXTRACAO=gemini-3.5-flash-lite    # leitura factual de páginas pesquisáveis
GEMINI_MODEL_OCR=gemini-3.6-flash              # leitura visual de páginas escaneadas e contratos
GEMINI_MODEL_RELATORIO=gemini-3.6-flash        # consolidação e elaboração da análise jurídica
GEMINI_MODEL_ESTRUTURADO=gemini-3.5-flash-lite # dados estruturados (JSON) de dossiês e cronologias
GEMINI_MODEL_QA=gemini-3.6-flash               # perguntas sobre relatórios já consolidados
GEMINI_MODEL_TIMELINE=gemini-2.5-pro           # leitura conjunta dos atos societários
```

A Timeline Societária usa modelo próprio porque submete todos os atos em uma chamada só, e o
default dela é de outra família (`gemini-2.5-pro`). Se o modelo configurado não existir na chave em
uso, o módulo interrompe com a lista dos modelos disponíveis, em vez de falhar em silêncio.

Em páginas com camada textual confiável, `GEMINI_MODEL_EXTRACAO` faz a leitura factual em alto
volume. Se o detector local encontrar qualquer página digitalizada no trecho, `GEMINI_MODEL_OCR`
assume a leitura visual desse trecho — seu default é o mesmo modelo forte da consolidação.

Processo e RJ carregam em cada extração o nome do arquivo e a página absoluta do PDF original. Os
marcadores internos de parte/chunk não podem aparecer no relatório: toda referência final combina o
identificador real do tribunal com a página absoluta, e o sufixo `do pdf ...` só é usado quando a
análise contém mais de um PDF. Fontes de RJ acima do limite de contexto são cobertas em lotes
intermediários, sem descartar o final do processo. Se qualquer chunk falhar, o relatório parcial é
bloqueado em vez de ser apresentado como completo.

### Chaves de acesso, failover e timeout
As chaves ficam como **secrets**: `GEMINI_API_KEY_1`, `GEMINI_API_KEY_2`, … lidas nessa ordem, sem
pular número (`GEMINI_API_KEY`, sem sufixo, é aceita como primeira). Podem pertencer a projetos ou
contas Google diferentes, cada uma com sua cota, o que permite distribuir carga.

Quando uma chamada falha por motivo que pode variar entre projetos — autenticação, permissão,
indisponibilidade do modelo, esgotamento de cota, limite de requisições — o sistema tenta a próxima
chave e preserva o fluxo. Cada chamada tem espera máxima de **10 minutos** (`GEMINI_TIMEOUT_MS`,
default 600000); passado o intervalo, a tentativa é encerrada para a análise não ficar travada.
Conforme a classificação do erro, a operação é repetida ou outra chave é acionada — e o log registra
qual credencial recusou e por quê, distinguindo falha de transporte de problema no documento.

### Limites de processamento
Também já são os defaults; só precisam ser cadastrados como Variables se houver necessidade de
alterá-los:

```text
MAX_ACTIVE_ANALYSES=4                 # análises ativas em Processos e RJ; as demais vão para a fila
TOTAL_EXTRACTION_WORKERS=6            # workers de extração compartilhados
PDF_PREPARATION_CONCURRENCY=2         # preparações de PDF simultâneas
PDF_CHUNK_MAX_MB=45                   # teto de tamanho por parte
LIMITE_PROCESSAMENTO_PESADO_PDF=6     # chunks em compressão/rasterização ao mesmo tempo no container
MATRICULAS_MAX_WORKERS=3              # matrículas processadas em paralelo
ANALYSIS_CACHE_RETENTION_DAYS=7       # retenção do cache de análises
RJ_CACHE_RETENCAO_DIAS=7              # retenção do cache de chunks de RJ
RJ_CACHE_RETENCAO_DIAS_CONCLUIDO=2    # idem, para análise já concluída
```

`LIMITE_PROCESSAMENTO_PESADO_PDF` limita só a etapa de compressão/rasterização, que é o gargalo real
de CPU e memória — a espera pela resposta do Gemini, que é I/O, não passa por essa trava.

O cache e o último estágio ficam em `/data/analysis_runtime` quando o Space tem storage persistente,
ou em `resultados/analysis_runtime` no ambiente local. `ANALYSIS_RUNTIME_DIR` e `RJ_CACHE_DIR`
sobrescrevem esses caminhos, e `PORT` (default 7860) muda a porta em que o app sobe.

## Desenvolvimento

### Rodar local
```bash
python app.py            # http://localhost:7860
```

### Testes
```bash
.venv_test/Scripts/python.exe -m pytest tests/ -q
```

### Ambiente igual ao do Space
O Space roda a versão de Gradio declarada em `sdk_version` acima. Para reproduzir localmente e pegar
antes do deploy os erros que só aparecem lá:

```bash
uv venv .venv_space --python 3.11
uv pip install --python .venv_space/Scripts/python.exe -r requirements.txt "gradio==5.29.1" pytest
```

### Publicação
1. Trabalhar em branch, nunca direto na `main`
2. `pytest` verde e app conferido localmente
3. Push da branch: o workflow `testes-e-space-teste.yml` roda os testes e publica no Space de
   **teste**
4. Conferência no Space de teste
5. Só então merge na `main`, que publica em **produção** pelo `sync.yml`

## Estado atual

- **Em produção:** Processos, Recuperação Judicial, Matrículas, Timeline Societária, Coleta de
  Informações
- **Em teste:** Excel consolidado de credores da RJ
- **Em aperfeiçoamento:** Cronologia Processual
- **Em desenvolvimento inicial:** Monitor de Safra — acompanhamento de áreas rurais ligadas a
  devedores por séries temporais de índices de vegetação e sensoriamento remoto, para estimar plantio,
  desenvolvimento e colheita e subsidiar o momento de requerer constrição sobre a produção (penhora
  de safra)
