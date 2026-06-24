# Apps Script - Gestao Financeira Pessoal

Este projeto transforma a planilha em uma base de dados no Google Sheets e usa HTML como tela de uso diario.

## Arquivos

- `Code.gs`: backend Apps Script, regras de negocio, criacao das abas e calculos.
- `Index.html`: estrutura da tela web.
- `Styles.html`: estilos responsivos e navegacao mobile-first.
- `Scripts.html`: interacoes, renderizacao, filtros e chamadas ao Apps Script.
- `appsscript.json`: configuracao do projeto Apps Script.

## Como instalar

1. Abra o Google Drive.
2. Importe `Gestao_Despesas_Receitas_Pessoais.xlsx`.
3. Abra o arquivo importado como Google Sheets.
4. No menu, acesse `Extensoes > Apps Script`.
5. Crie/cole os arquivos:
   - `Code.gs`
   - `Index.html`
   - `Styles.html`
   - `Scripts.html`
   - `appsscript.json`
6. Se o Apps Script foi criado dentro da propria planilha, rode `setupPlanilha`.
7. Se o Apps Script foi criado como projeto separado, copie o ID da URL do Google Sheets e escolha uma das formas:
   - Recomendada para uso fixo: preencha `SPREADSHEET_ID_FIXO` no topo do `Code.gs`.
   - Alternativa sem editar o topo do arquivo: rode `configurarSpreadsheetId("ID_DA_PLANILHA")` uma vez.
   Depois rode `setupPlanilha`.
8. Autorize o acesso solicitado pelo Google.
9. Para conferir a ligacao com a planilha, rode `diagnosticarProjeto`.
10. Clique em `Implantar > Nova implantacao > App da Web`.
11. Configure:
   - Executar como: `Eu`.
   - Quem pode acessar: `Somente eu` ou conforme sua necessidade.
12. Abra a URL gerada.

## Uso no celular

- A interface usa navegacao inferior com telas separadas para inicio, lancamentos, metas e cadastro.
- Os lancamentos sao exibidos como cartoes no celular.
- A listagem carrega 30 registros por vez; use `Carregar mais` para consultar registros antigos.
- A busca aguarda alguns milissegundos antes de consultar o servidor, reduzindo chamadas durante a digitacao.
- Ao alterar qualquer arquivo, crie uma nova versao da implantacao para atualizar a URL `/exec`.

## Orientacoes para a planilha

A planilha pode estar em branco. A funcao `setupPlanilha` cria e ajusta estas abas:

- `Lancamentos`: armazenamento principal de receitas e despesas.
- `Categorias`: listas de tipos, categorias, formas de pagamento, contas, status e subcategorias.
- `Resumo Mensal`: consolidado por mes.
- `Dashboard`: indicadores principais.
- `Metas`: limites mensais por categoria.
- `Ajuda`: instrucoes dentro da propria planilha.

Nao apague a primeira linha de nenhuma aba, pois ela contem os cabecalhos usados pelo Apps Script.

Se a aba `Lancamentos` estiver no modelo antigo da planilha local, com cabecalho iniciando em `Data, Tipo, Categoria`, o script migra esses registros automaticamente para o novo modelo com `ID`, datas de criacao e datas de atualizacao.

## Regra de negocio implementada

- Todo lancamento tem data, tipo, categoria, descricao, forma, conta, valor e status.
- O valor deve ser positivo.
- O campo `Tipo` define se o valor entra como receita ou despesa.
- Lancamentos com status `Previsto` ficam armazenados, mas nao entram como realizados no dashboard principal.
- O campo `Mes` e gerado automaticamente no formato `AAAA-MM`.
- O dashboard calcula receitas, despesas, saldo, taxa de poupanca e maior categoria de despesa.
- A aba `Metas` compara o limite mensal de cada categoria contra os gastos realizados no mes atual.

## Melhorias sugeridas na planilha

Depois de rodar `setupPlanilha`, ajuste a aba `Categorias`:

- Edite as categorias de despesa e receita conforme sua realidade.
- Edite as contas, por exemplo banco principal, carteira, cartao de credito e poupanca.
- Edite formas de pagamento.
- Mantenha categorias sem acento se quiser evitar problemas em filtros antigos do Excel/Sheets.

Para controle mais avancado, adicione novas categorias na aba `Metas` e preencha `Limite Mensal`.
