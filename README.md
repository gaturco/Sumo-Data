# Sumo API Integration (Planilha + Google Apps Script)

Este repositório contém um script Google Apps Script (`sumo_data.gs`) que integra uma planilha do Google Sheets com uma API de resultados de sumô. O objetivo é automatizar a criação de abas por *basho* (torneio), atualizar a lista de *rikishi* (lutadores) e preencher os resultados por dia.

Planilha de referência: https://docs.google.com/spreadsheets/d/1IREDADljXm8x8FOpxU9m3wxKD0hompdTT_bc0HVi2po/edit?gid=1144630652

## Conteúdo

- `sumo_data.gs` — script principal em Google Apps Script
- `README.md` — este arquivo

## Visão geral

Principais funções:

- `updateSumoResults()` — pergunta o dia (1–15), consulta a API de `torikumi` e escreve os resultados na aba ativa. Adiciona uma coluna `Dia X` e marca `W`, `L` ou `VS` para confrontos diretos.
- `updateRikishiList()` — solicita o Basho (mm/yyyy), busca o `banzuke` e atualiza a aba `Pepperoni` adicionando novos rikishi e atualizando a coluna de bashos (coluna C).
- `createNewBashoSheet()` — copia uma aba modelo (`TEMPLATE_SHEET_NAME`), renomeia para o Basho (mm/yyyy) e preenche a Coluna A com rikishi filtrados pela aba `Pepperoni` marcados com `S`.

### Configurações principais no script

- `DIVISION` — divisão (ex.: `Makuuchi`)
- `RIKISHI_COLUMN_INDEX` — índice da coluna onde estão os nomes (padrão `1` = Coluna A)
- `PEPPERONI_COLUMN` — coluna que marca participação (padrão `2` = Coluna B)
- `BASHO_LIST_COLUMN` — coluna que contém a lista de bashos do rikishi (padrão `3` = Coluna C)
- `TEMPLATE_SHEET_NAME` — nome da aba modelo usada para criar novas abas (ex.: `11/2025`)
- `FIRST_DATA_ROW` — primeira linha de dados (padrão `2`, isto é, A2)
- `WIN_RATE_LABEL` — rótulo de linha final para Win Rate

## Layout esperado da planilha (aba modelo e `Pepperoni`)

- **Linha 1:** cabeçalho (nomes das colunas)
- **Coluna A:** `Rikishi` — nomes dos competidores (começando em A2)
- **Coluna B:** marca `S` quando o rikishi participa de um basho específico (usado por `createNewBashoSheet`)
- **Coluna C:** lista de bashos que o rikishi participou no formato `mm/yyyy,mm/yyyy` (ex.: `11/2025,01/2026`)
- **Linha final:** a última linha abaixo dos rikishi é usada para `Win Rate`

## Uso rápido (manual)

1. Abra a planilha no Google Sheets.
2. Abra o Editor de Scripts: **Extensões → Apps Script**.
3. Cole/atualize o conteúdo do `sumo_data.gs` no editor e salve.
4. Execute qualquer função (ex.: `updateRikishiList`) pelo menu de funções do Apps Script — a primeira execução pedirá permissões (`UrlFetchApp`, acesso à planilha etc.).
5. (Opcional) Crie um botão na planilha e atribua a função desejada: **Inserir → Desenho** → desenhe o botão → selecione o desenho → **Atribuir script**.

## Exemplos de fluxo

- **Criar uma aba nova para o Basho `11/2025`:**
	1. Execute `createNewBashoSheet()` no Apps Script (ou atribua a um botão).
	2. Insira `11/2025` quando solicitado. O script copiará a aba definida em `TEMPLATE_SHEET_NAME` e preencherá a Coluna A com os rikishi filtrados da aba `Pepperoni`.

- **Atualizar resultados do Dia 5 em uma aba de basho:**
	1. Abra a aba do basho (nome no formato `mm/yyyy`).
	2. Execute `updateSumoResults()` e informe o dia `5` quando solicitado.
	3. O script chamará `https://sumo-api.com/api/basho/{yyyyMM}/torikumi/{DIVISION}/{day}` e preencherá a coluna `Dia 5` com `W/L/VS`.

## Integração com a API

O script usa `UrlFetchApp.fetch(apiUrl)` para acessar os endpoints:

- `/api/basho/{yyyyMM}/banzuke/{DIVISION}` — lista de rikishi (east/west)
- `/api/basho/{yyyyMM}/torikumi/{DIVISION}/{day}` — confrontos do dia

Certifique-se de que a API pública esteja disponível e que o domínio `sumo-api.com` esteja acessível ao Apps Script.

## Permissões necessárias

- `UrlFetchApp` (acesso externo à API)
- Acesso à planilha para leitura/gravação de ranges e abas

## Deploy / Dicas de implantação

- Ferramenta recomendada: Editor Apps Script (embed no próprio Google Sheets). Não é necessário criar um repositório separado para rodar o script.
- Na primeira execução, revise e autorize as permissões solicitadas.
- Para versionamento local e integração com Git, considere usar o Google `clasp` (https://github.com/google/clasp) para sincronizar o projeto Apps Script com um repositório Git.

## Boas práticas / Troubleshooting

- Nomes de abas precisam estar no formato `mm/yyyy` para que `getBashoIdFromSheetName()` funcione corretamente.
- Se alguma chamada de API falhar, abra o Apps Script e verifique o erro exibido pelo editor ou nos logs.
- Se o script não encontrar rikishi esperados, verifique espaços em branco e diferenças de capitalização — o script faz comparações exatas de string.
- Para versionamento local, use `clasp` e autentique com sua conta Google.

## Seções futuras sugeridas

- Exemplo de exportação CSV do layout
- Testes unitários simulando respostas da API
- Automação: gatilhos temporais no Apps Script para atualizar resultados automaticamente

## Autor

- Gabriel Turco (autor do `sumo_data.gs`)
