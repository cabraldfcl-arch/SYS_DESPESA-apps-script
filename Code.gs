const SPREADSHEET_ID_FIXO = '1JpuZHd0LjcAqO-oTi6kD9aMxr1_mwQjmtvqQ7kvzLaU';
let SPREADSHEET_CACHE = null;

const CONFIG = {
  spreadsheetId: SPREADSHEET_ID_FIXO,
  sheets: {
    lancamentos: 'Lancamentos',
    categorias: 'Categorias',
    resumo: 'Resumo Mensal',
    dashboard: 'Dashboard',
    metas: 'Metas',
    ajuda: 'Ajuda'
  },
  headers: {
    lancamentos: [
      'ID',
      'Data',
      'Tipo',
      'Categoria',
      'Subcategoria',
      'Descricao',
      'Forma',
      'Conta',
      'Valor',
      'Mes',
      'Status',
      'Observacoes',
      'Criado em',
      'Atualizado em'
    ],
    categorias: [
      'Tipos',
      '',
      'Categorias de Despesa',
      '',
      'Categorias de Receita',
      '',
      'Todas as Categorias',
      '',
      'Formas',
      '',
      'Contas',
      '',
      'Status'
    ],
    resumo: [
      'Mes',
      'Receitas',
      'Despesas',
      'Saldo',
      'Taxa de Poupanca',
      'Categoria Despesa',
      'Total',
      'Participacao'
    ],
    metas: ['Categoria', 'Limite Mensal', 'Gasto Mes Atual', 'Diferenca', 'Status']
  },
  defaults: {
    tipos: ['Receita', 'Despesa'],
    despesas: ['Moradia', 'Alimentacao', 'Transporte', 'Saude', 'Educacao', 'Lazer', 'Assinaturas', 'Mercado', 'Cartao', 'Outros'],
    receitas: ['Salario', 'Freelance', 'Investimentos', 'Reembolsos', 'Vendas', 'Presente', 'Outros'],
    formas: ['Pix', 'Cartao Debito', 'Cartao Credito', 'Dinheiro', 'Boleto', 'Transferencia', 'Debito Automatico', 'Outros'],
    contas: ['Conta Corrente', 'Carteira', 'Poupanca', 'Cartao 1', 'Cartao 2', 'Investimentos'],
    status: ['Pago', 'Pendente', 'Previsto'],
    subcategorias: ['Aluguel', 'Condominio', 'Energia', 'Agua', 'Internet', 'Mercado', 'Combustivel', 'Medicamentos', 'Curso', 'Restaurante']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setTitle('Gestao Financeira Pessoal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupPlanilha() {
  const ss = getSpreadsheet_();
  setupCategorias_(ss);
  setupLancamentos_(ss);
  setupResumo_(ss);
  setupDashboard_(ss);
  setupMetas_(ss);
  setupAjuda_(ss);
  recalcularResumoMetasDashboard_();
  return { ok: true };
}

function configurarSpreadsheetId(id) {
  const spreadsheetId = String(id || '').trim();
  if (!spreadsheetId) {
    throw new Error('Informe o ID da planilha do Google Sheets.');
  }
  SpreadsheetApp.openById(spreadsheetId);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
  SPREADSHEET_CACHE = null;
  return { ok: true, spreadsheetId };
}

function limparSpreadsheetId() {
  PropertiesService.getScriptProperties().deleteProperty('SPREADSHEET_ID');
  SPREADSHEET_CACHE = null;
  return { ok: true };
}

function diagnosticarProjeto() {
  const result = {
    spreadsheetConfigurada: false,
    spreadsheetId: '',
    spreadsheetNome: '',
    abas: [],
    webApp: true
  };

  const ss = getSpreadsheet_();
  result.spreadsheetConfigurada = true;
  result.spreadsheetId = ss.getId();
  result.spreadsheetNome = ss.getName();
  result.abas = ss.getSheets().map(sheet => sheet.getName());
  return result;
}

function getBootstrapData(filtros) {
  filtros = filtros || {};
  if (!filtros.limite) filtros.limite = 30;
  const lancamentos = getLancamentosNormalizados_();
  return {
    listas: getListas(lancamentos),
    dashboard: getDashboardSemSetup_(lancamentos),
    resumo: getResumoMensalSemSetup_(lancamentos),
    metas: getMetasSemSetup_(lancamentos),
    lancamentos: paginarLancamentos_(filtros, lancamentos)
  };
}

function getListas(lancamentos) {
  const sheet = getSheet_(CONFIG.sheets.categorias);
  const cols = getCategoriaColumns_(sheet);
  const values = sheet.getDataRange().getValues();
  return {
    tipos: controlledList_(readColumnFromValues_(values, cols.tipos), CONFIG.defaults.tipos),
    categoriasDespesa: listOrDefault_(readColumnFromValues_(values, cols.despesas), CONFIG.defaults.despesas),
    categoriasReceita: listOrDefault_(readColumnFromValues_(values, cols.receitas), CONFIG.defaults.receitas),
    categorias: listOrDefault_(readColumnFromValues_(values, cols.todas), unique_(CONFIG.defaults.receitas.concat(CONFIG.defaults.despesas))),
    formas: controlledList_(readColumnFromValues_(values, cols.formas), CONFIG.defaults.formas),
    contas: controlledList_(readColumnFromValues_(values, cols.contas), CONFIG.defaults.contas),
    status: controlledList_(readColumnFromValues_(values, cols.status), CONFIG.defaults.status),
    subcategorias: listOrDefault_(getSubcategorias_(lancamentos), CONFIG.defaults.subcategorias)
  };
}

function salvarLancamento(payload) {
  const data = normalizarLancamento_(payload);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_(CONFIG.sheets.lancamentos);
    const now = new Date();
    const id = data.id || Utilities.getUuid();
    const row = [
      id,
      data.data,
      data.tipo,
      data.categoria,
      data.subcategoria,
      data.descricao,
      data.forma,
      data.conta,
      data.valor,
      mesKey_(data.data),
      data.status,
      data.observacoes,
      now,
      now
    ];

    if (data.id) {
      const rowIndex = findRowById_(sheet, data.id);
      if (!rowIndex) {
        throw new Error('Lancamento nao encontrado para atualizacao.');
      }
      row[12] = sheet.getRange(rowIndex, 13).getValue() || now;
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    }

    return {
      ok: true,
      lancamento: normalizarSaidaLancamento_(rowToObject_(CONFIG.headers.lancamentos, row))
    };
  } finally {
    lock.releaseLock();
  }
}

function excluirLancamento(id) {
  if (!id) {
    throw new Error('Informe o ID do lancamento.');
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_(CONFIG.sheets.lancamentos);
    const rowIndex = findRowById_(sheet, id);
    if (!rowIndex) {
      throw new Error('Lancamento nao encontrado.');
    }
    sheet.deleteRow(rowIndex);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function listarLancamentos(filtros) {
  return paginarLancamentos_(filtros);
}

function listarLancamentosSemSetup_(filtros, lancamentos) {
  filtros = filtros || {};
  const values = lancamentos || getLancamentosNormalizados_();
  const mes = filtros.mes || '';
  const tipo = filtros.tipo || '';
  const texto = String(filtros.texto || '').toLowerCase();

  return values
    .filter(item => !mes || item.mes === mes)
    .filter(item => !tipo || item.tipo === tipo)
    .filter(item => {
      if (!texto) return true;
      return [item.categoria, item.subcategoria, item.descricao, item.conta, item.status, item.observacoes]
        .join(' ')
        .toLowerCase()
        .indexOf(texto) >= 0;
    })
    .sort((a, b) => new Date(b.data).getTime() - new Date(a.data).getTime());
}

function paginarLancamentos_(filtros, lancamentos) {
  filtros = filtros || {};
  const limite = Math.min(Math.max(Number(filtros.limite || 30), 1), 100);
  const offset = Math.max(Number(filtros.offset || 0), 0);
  const filtrados = listarLancamentosSemSetup_(filtros, lancamentos);
  const items = filtrados.slice(offset, offset + limite);

  return {
    items,
    total: filtrados.length,
    offset,
    limite,
    hasMore: offset + items.length < filtrados.length,
    nextOffset: offset + items.length
  };
}

function getDashboard() {
  setupPlanilha();
  return getDashboardSemSetup_();
}

function getResumoMensal() {
  setupPlanilha();
  return getResumoMensalSemSetup_();
}

function getDashboardSemSetup_(lancamentos) {
  lancamentos = lancamentos || getLancamentosNormalizados_();
  const hoje = new Date();
  const mesAtual = mesKey_(hoje);
  const anoAtual = hoje.getFullYear();
  const pagos = lancamentos.filter(item => item.status !== 'Previsto');
  const doAno = pagos.filter(item => parseDate_(item.data).getFullYear() === anoAtual);
  const doMes = pagos.filter(item => item.mes === mesAtual);

  const receitasAno = somaNormalizada_(doAno, 'Receita');
  const despesasAno = somaNormalizada_(doAno, 'Despesa');
  const receitasMes = somaNormalizada_(doMes, 'Receita');
  const despesasMes = somaNormalizada_(doMes, 'Despesa');
  const saldoMes = receitasMes - despesasMes;
  const taxaPoupanca = receitasMes > 0 ? saldoMes / receitasMes : 0;
  const porCategoria = agruparDespesasPorCategoriaNormalizadas_(doMes);
  const maiorCategoria = porCategoria.length ? porCategoria[0] : { categoria: '-', total: 0 };

  return {
    mesAtual,
    receitasAno,
    despesasAno,
    saldoAno: receitasAno - despesasAno,
    receitasMes,
    despesasMes,
    saldoMes,
    taxaPoupanca,
    maiorCategoria: maiorCategoria.categoria,
    maiorCategoriaValor: maiorCategoria.total,
    totalLancamentos: lancamentos.length,
    porCategoria
  };
}

function getMetas() {
  setupPlanilha();
  return getMetasSemSetup_();
}

function getMetasSemSetup_(lancamentos) {
  const mesAtual = mesKey_(new Date());
  const despesas = (lancamentos || getLancamentosNormalizados_())
    .filter(item => item.tipo === 'Despesa' && item.mes === mesAtual && item.status !== 'Previsto');

  return sheetToObjects_(getSheet_(CONFIG.sheets.metas)).map(item => {
    const categoria = item.Categoria;
    const limiteMensal = number_(item['Limite Mensal']);
    const gastoMesAtual = despesas
      .filter(lancamento => lancamento.categoria === categoria)
      .reduce((total, lancamento) => total + number_(lancamento.valor), 0);
    const diferenca = limiteMensal - gastoMesAtual;
    let status = 'Sem limite';
    if (limiteMensal > 0 && diferenca >= 0) status = 'Dentro da meta';
    if (limiteMensal > 0 && diferenca < 0) status = 'Acima da meta';

    return { categoria, limiteMensal, gastoMesAtual, diferenca, status };
  });
}

function salvarMeta(payload) {
  const categoria = String(payload && payload.categoria || '').trim();
  const limite = number_(payload && payload.limiteMensal);
  if (!categoria) {
    throw new Error('Informe a categoria da meta.');
  }
  if (limite < 0) {
    throw new Error('A meta nao pode ser negativa.');
  }

  const sheet = getSheet_(CONFIG.sheets.metas);
  const rows = sheet.getDataRange().getValues();
  let rowIndex = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === categoria) {
      rowIndex = i + 1;
      break;
    }
  }

  if (!rowIndex) {
    rowIndex = sheet.getLastRow() + 1;
    sheet.getRange(rowIndex, 1).setValue(categoria);
  }

  sheet.getRange(rowIndex, 2).setValue(limite);
  return { ok: true, categoria, limiteMensal: limite };
}

function recalcularResumoMetasDashboard_() {
  recalcularResumo_();
  recalcularMetas_();
  escreverDashboard_();
}

function recalcularDadosAuxiliares() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    recalcularResumoMetasDashboard_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function setupLancamentos_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.lancamentos);
  migrateLegacyLancamentos_(sheet);
  ensureHeaders_(sheet, CONFIG.headers.lancamentos);
  ensureLancamentoMesKeys_(sheet);
  aplicarFormatacaoLancamentos_(sheet);
  clearLancamentoValidations_(sheet);
  applyValidation_(sheet, 3, CONFIG.defaults.tipos);
  applyValidation_(sheet, 4, CONFIG.defaults.despesas.concat(CONFIG.defaults.receitas));
  applyValidation_(sheet, 5, CONFIG.defaults.subcategorias);
  applyValidation_(sheet, 7, CONFIG.defaults.formas);
  applyValidation_(sheet, 8, CONFIG.defaults.contas);
  applyValidation_(sheet, 11, CONFIG.defaults.status);
  sheet.setFrozenRows(1);
}

function migrateLegacyLancamentos_(sheet) {
  if (sheet.getLastRow() < 1) return;

  const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  const legacyHeader = firstRow[0] === 'Data' && firstRow[1] === 'Tipo' && firstRow[2] === 'Categoria';
  const alreadyCurrent = firstRow[0] === 'ID' && firstRow[1] === 'Data';
  if (!legacyHeader || alreadyCurrent) return;

  const values = sheet.getDataRange().getValues();
  const now = new Date();
  const rows = values.slice(1)
    .filter(row => row[0] && row[0] !== 0 && row[1] && row[2] && row[7])
    .map(row => {
      const date = parseDate_(row[0]);
      return [
        Utilities.getUuid(),
        date,
        row[1],
        row[2],
        row[3],
        row[4],
        row[5],
        row[6],
        number_(row[7]),
        isMesKey_(row[8]) ? row[8] : mesKey_(date),
        row[9] || 'Pago',
        row[10] || '',
        now,
        now
      ];
    });

  rewriteBody_(sheet, CONFIG.headers.lancamentos, rows);
}

function ensureLancamentoMesKeys_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, CONFIG.headers.lancamentos.length);
  const values = range.getValues();
  let changed = false;
  values.forEach(row => {
    const data = parseDate_(row[1]);
    if (!data || isNaN(data.getTime())) return;
    if (isMesKey_(row[9])) return;
    row[9] = mesKey_(data);
    changed = true;
  });
  if (changed) {
    range.setValues(values);
  }
}

function setupCategorias_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.categorias);
  migrateCategorias_(sheet);
  ensureHeaders_(sheet, CONFIG.headers.categorias);
  const cols = getCategoriaColumns_(sheet);
  seedColumn_(sheet, cols.tipos, CONFIG.defaults.tipos);
  seedColumn_(sheet, cols.despesas, CONFIG.defaults.despesas);
  seedColumn_(sheet, cols.receitas, CONFIG.defaults.receitas);
  seedColumn_(sheet, cols.todas, unique_(CONFIG.defaults.receitas.concat(CONFIG.defaults.despesas)));
  replaceColumn_(sheet, cols.tipos, CONFIG.defaults.tipos);
  replaceColumn_(sheet, cols.formas, CONFIG.defaults.formas);
  replaceColumn_(sheet, cols.contas, CONFIG.defaults.contas);
  replaceColumn_(sheet, cols.status, CONFIG.defaults.status);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, CONFIG.headers.categorias.length);
}

function migrateCategorias_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(value => String(value || '').trim());
  const compactHeader = firstRow[0] === 'Tipos'
    && firstRow[1] === 'Categorias de Despesa'
    && firstRow[2] === 'Categorias de Receita';
  if (!compactHeader) return;

  const values = {
    tipos: readColumn_(sheet, 1),
    despesas: readColumn_(sheet, 2),
    receitas: readColumn_(sheet, 3),
    todas: readColumn_(sheet, 4),
    formas: readColumn_(sheet, 5),
    contas: readColumn_(sheet, 6),
    status: readColumn_(sheet, 7)
  };

  sheet.clearContents();
  sheet.getRange(1, 1, 1, CONFIG.headers.categorias.length).setValues([CONFIG.headers.categorias]);
  writeColumn_(sheet, 1, values.tipos);
  writeColumn_(sheet, 3, values.despesas);
  writeColumn_(sheet, 5, values.receitas);
  writeColumn_(sheet, 7, values.todas);
  writeColumn_(sheet, 9, values.formas);
  writeColumn_(sheet, 11, values.contas);
  writeColumn_(sheet, 13, values.status);
}

function setupResumo_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.resumo);
  ensureHeaders_(sheet, CONFIG.headers.resumo);
  if (sheet.getLastRow() < 2) {
    const ano = new Date().getFullYear();
    const rows = [];
    for (let m = 1; m <= 12; m++) {
      rows.push([`${ano}-${String(m).padStart(2, '0')}`, 0, 0, 0, 0, '', 0, 0]);
    }
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.getRange('B:D').setNumberFormat('R$ #,##0.00');
  sheet.getRange('E:E').setNumberFormat('0.00%');
}

function setupDashboard_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.dashboard);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Dashboard Financeiro Pessoal');
    sheet.getRange(3, 1, 9, 2).setValues([
      ['Mes atual', ''],
      ['Receitas no mes', 0],
      ['Despesas no mes', 0],
      ['Saldo no mes', 0],
      ['Taxa de poupanca', 0],
      ['Receitas no ano', 0],
      ['Despesas no ano', 0],
      ['Saldo no ano', 0],
      ['Maior categoria de despesa', '']
    ]);
  }
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('B4:B10').setNumberFormat('R$ #,##0.00');
  sheet.getRange('B7').setNumberFormat('0.00%');
}

function setupMetas_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.metas);
  ensureHeaders_(sheet, CONFIG.headers.metas);
  if (sheet.getLastRow() < 2) {
    const rows = CONFIG.defaults.despesas.map(categoria => [categoria, 0, 0, 0, 'Sem limite']);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.getRange('B:D').setNumberFormat('R$ #,##0.00');
}

function setupAjuda_(ss) {
  const sheet = ensureSheet_(ss, CONFIG.sheets.ajuda);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 9, 1).setValues([
      ['Como usar'],
      ['1. Rode setupPlanilha uma vez depois de importar a planilha para o Google Sheets.'],
      ['2. Use a tela HTML para cadastrar receitas e despesas.'],
      ['3. A aba Lancamentos armazena os dados. Nao apague a linha de cabecalho.'],
      ['4. Edite listas na aba Categorias para adaptar categorias, contas e formas.'],
      ['5. Defina limites na aba Metas para controlar gastos do mes atual.'],
      ['6. O Resumo Mensal e o Dashboard sao recalculados pelo Apps Script.'],
      ['7. Valores devem ser positivos; o campo Tipo define Receita ou Despesa.'],
      ['8. Status Previsto nao entra como realizado no dashboard principal.']
    ]);
  }
  sheet.autoResizeColumn(1);
}

function recalcularResumo_() {
  const sheet = getSheet_(CONFIG.sheets.resumo);
  const rows = getResumoMensalSemSetup_().map(item => [
    item.mes,
    item.receitas,
    item.despesas,
    item.saldo,
    item.taxaPoupanca,
    '',
    0,
    0
  ]);
  rewriteBody_(sheet, CONFIG.headers.resumo, rows);
  sheet.getRange('B:D').setNumberFormat('R$ #,##0.00');
  sheet.getRange('E:E').setNumberFormat('0.00%');
}

function recalcularMetas_() {
  const sheet = getSheet_(CONFIG.sheets.metas);
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return;

  const mesAtual = mesKey_(new Date());
  const lancamentos = getLancamentosNormalizados_()
    .filter(item => item.tipo === 'Despesa' && item.mes === mesAtual && item.status !== 'Previsto');

  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const categoria = rows[i][0];
    if (!categoria) continue;
    const limite = number_(rows[i][1]);
    const gasto = lancamentos
      .filter(item => item.categoria === categoria)
      .reduce((sum, item) => sum + number_(item.valor), 0);
    const diferenca = limite - gasto;
    let status = 'Sem limite';
    if (limite > 0 && diferenca >= 0) status = 'Dentro da meta';
    if (limite > 0 && diferenca < 0) status = 'Acima da meta';
    result.push([categoria, limite, gasto, diferenca, status]);
  }

  rewriteBody_(sheet, CONFIG.headers.metas, result);
  sheet.getRange('B:D').setNumberFormat('R$ #,##0.00');
}

function escreverDashboard_() {
  const sheet = getSheet_(CONFIG.sheets.dashboard);
  const dashboard = getDashboardSemSetup_();
  sheet.getRange(3, 1, 9, 2).setValues([
    ['Mes atual', dashboard.mesAtual],
    ['Receitas no mes', dashboard.receitasMes],
    ['Despesas no mes', dashboard.despesasMes],
    ['Saldo no mes', dashboard.saldoMes],
    ['Taxa de poupanca', dashboard.taxaPoupanca],
    ['Receitas no ano', dashboard.receitasAno],
    ['Despesas no ano', dashboard.despesasAno],
    ['Saldo no ano', dashboard.saldoAno],
    ['Maior categoria de despesa', `${dashboard.maiorCategoria} - ${formatCurrency_(dashboard.maiorCategoriaValor)}`]
  ]);
}

function normalizarLancamento_(payload) {
  payload = payload || {};
  const data = parseDate_(payload.data);
  const tipo = String(payload.tipo || '').trim();
  const categoria = String(payload.categoria || '').trim();
  const descricao = String(payload.descricao || '').trim();
  const valor = number_(payload.valor);

  if (!data || isNaN(data.getTime())) throw new Error('Informe uma data valida.');
  if (CONFIG.defaults.tipos.indexOf(tipo) < 0) throw new Error('Tipo deve ser Receita ou Despesa.');
  if (!categoria) throw new Error('Informe a categoria.');
  if (!descricao) throw new Error('Informe a descricao.');
  if (valor <= 0) throw new Error('Informe um valor positivo.');

  return {
    id: payload.id || '',
    data,
    tipo,
    categoria,
    subcategoria: String(payload.subcategoria || '').trim(),
    descricao,
    forma: String(payload.forma || 'Pix').trim(),
    conta: String(payload.conta || 'Conta Corrente').trim(),
    valor,
    status: String(payload.status || 'Pago').trim(),
    observacoes: String(payload.observacoes || '').trim()
  };
}

function ensureSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error(`Aba nao encontrada: ${name}. Rode setupPlanilha.`);
  return sheet;
}

function getSpreadsheet_() {
  if (SPREADSHEET_CACHE) {
    return SPREADSHEET_CACHE;
  }

  const configuredId = CONFIG.spreadsheetId || PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (configuredId) {
    SPREADSHEET_CACHE = SpreadsheetApp.openById(configuredId);
    return SPREADSHEET_CACHE;
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    SPREADSHEET_CACHE = active;
    return SPREADSHEET_CACHE;
  }

  throw new Error('Nenhuma planilha ativa encontrada. Vincule o script a uma planilha ou rode configurarSpreadsheetId("ID_DA_PLANILHA").');
}

function ensureHeaders_(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const mustWrite = headers.some((header, index) => current[index] !== header);
  if (mustWrite) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.autoResizeColumns(1, headers.length);
}

function seedColumn_(sheet, col, values) {
  const current = readColumn_(sheet, col);
  if (current.length) return;
  sheet.getRange(2, col, values.length, 1).setValues(values.map(value => [value]));
}

function readColumn_(sheet, col) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues().flat();
  return values.map(value => String(value || '').trim()).filter(Boolean);
}

function readColumnFromValues_(values, col) {
  if (!col || !values || values.length < 2) return [];
  return unique_(values.slice(1)
    .map(row => String(row[col - 1] || '').trim())
    .filter(Boolean));
}

function writeColumn_(sheet, col, values) {
  if (!values || !values.length) return;
  sheet.getRange(2, col, values.length, 1).setValues(values.map(value => [value]));
}

function replaceColumn_(sheet, col, values) {
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, rowCount, 1).clearContent();
  writeColumn_(sheet, col, values);
}

function getCategoriaColumns_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), CONFIG.headers.categorias.length);
  const firstRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(value => String(value || '').trim());
  const spacedLayout = firstRow[2] === 'Categorias de Despesa'
    || firstRow[4] === 'Categorias de Receita'
    || firstRow[8] === 'Formas';

  if (spacedLayout) {
    return {
      tipos: 1,
      despesas: 3,
      receitas: 5,
      todas: 7,
      formas: 9,
      contas: 11,
      status: 13
    };
  }

  return {
    tipos: 1,
    despesas: 2,
    receitas: 3,
    todas: 4,
    formas: 5,
    contas: 6,
    status: 7
  };
}

function getSubcategorias_(lancamentos) {
  try {
    if (lancamentos) {
      return unique_(lancamentos.map(item => item.subcategoria).concat(CONFIG.defaults.subcategorias));
    }
    return unique_(readColumn_(getSheet_(CONFIG.sheets.lancamentos), 5).concat(CONFIG.defaults.subcategorias));
  } catch (error) {
    return CONFIG.defaults.subcategorias.slice();
  }
}

function listOrDefault_(values, defaults) {
  return values && values.length ? values : defaults.slice();
}

function controlledList_(values, defaults) {
  const allowed = unique_((values || []).filter(value => defaults.indexOf(value) >= 0));
  return allowed.length === defaults.length ? allowed : defaults.slice();
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(header => String(header || '').trim());
  return values.slice(1)
    .filter(row => row.some(value => value !== '' && value !== null))
    .map(row => {
      const item = {};
      headers.forEach((header, index) => item[header] = row[index]);
      return item;
    });
}

function rowToObject_(headers, row) {
  const item = {};
  headers.forEach((header, index) => item[header] = row[index]);
  return item;
}

function rewriteBody_(sheet, headers, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  ensureHeaders_(sheet, headers);
}

function applyValidation_(sheet, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(unique_(values), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function clearLancamentoValidations_(sheet) {
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 1, rowCount, CONFIG.headers.lancamentos.length).clearDataValidations();
}

function aplicarFormatacaoLancamentos_(sheet) {
  sheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
  sheet.getRange('I:I').setNumberFormat('R$ #,##0.00');
  sheet.getRange('J:J').setNumberFormat('@');
  sheet.getRange('M:N').setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(1, 1, 1, CONFIG.headers.lancamentos.length)
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), CONFIG.headers.lancamentos.length)
    .setVerticalAlignment('middle');
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).setHorizontalAlignment('center');
  sheet.getRange(2, 9, Math.max(sheet.getMaxRows() - 1, 1), 1).setHorizontalAlignment('right');
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 105);
  sheet.setColumnWidth(3, 105);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 220);
  sheet.setColumnWidth(7, 150);
  sheet.setColumnWidth(8, 145);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 90);
  sheet.setColumnWidth(11, 105);
  sheet.setColumnWidth(12, 260);
  sheet.setColumnWidths(13, 2, 150);
  hideLancamentoTechnicalColumns_(sheet);
  ensureFilter_(sheet, CONFIG.headers.lancamentos.length);
}

function hideLancamentoTechnicalColumns_(sheet) {
  sheet.hideColumns(1);
  sheet.hideColumns(10);
  sheet.hideColumns(13, 2);
}

function ensureFilter_(sheet, columnCount) {
  if (sheet.getFilter()) return;
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), columnCount).createFilter();
}

function findRowById_(sheet, id) {
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === id) return i + 2;
  }
  return 0;
}

function parseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  if (typeof value === 'number') return new Date(Math.round((value - 25569) * 86400 * 1000));
  const text = String(value || '').trim();
  if (!text) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    const parts = text.split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const parts = text.split('/').map(Number);
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(text);
}

function mesKey_(date) {
  const d = parseDate_(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function isMesKey_(value) {
  return /^\d{4}-\d{2}$/.test(String(value || '').trim());
}

function getLancamentosNormalizados_() {
  return sheetToObjects_(getSheet_(CONFIG.sheets.lancamentos))
    .map(normalizarSaidaLancamento_)
    .filter(Boolean);
}

function getResumoMensalSemSetup_(lancamentos) {
  lancamentos = (lancamentos || getLancamentosNormalizados_())
    .filter(item => item.status !== 'Previsto');
  const meses = unique_(lancamentos.map(item => item.mes).filter(Boolean)).sort();
  const ano = new Date().getFullYear();
  const mesesBase = [];
  for (let m = 1; m <= 12; m++) {
    mesesBase.push(`${ano}-${String(m).padStart(2, '0')}`);
  }

  return unique_(mesesBase.concat(meses)).sort().map(mes => {
    const itens = lancamentos.filter(item => item.mes === mes);
    const receitas = somaNormalizada_(itens, 'Receita');
    const despesas = somaNormalizada_(itens, 'Despesa');
    const saldo = receitas - despesas;
    return {
      mes,
      receitas,
      despesas,
      saldo,
      taxaPoupanca: receitas > 0 ? saldo / receitas : 0
    };
  });
}

function somaNormalizada_(items, tipo) {
  return items
    .filter(item => item.tipo === tipo)
    .reduce((sum, item) => sum + number_(item.valor), 0);
}

function agruparDespesasPorCategoriaNormalizadas_(items) {
  const map = {};
  items
    .filter(item => item.tipo === 'Despesa')
    .forEach(item => {
      const categoria = item.categoria || 'Sem categoria';
      map[categoria] = (map[categoria] || 0) + number_(item.valor);
    });
  return Object.keys(map)
    .map(categoria => ({ categoria, total: map[categoria] }))
    .sort((a, b) => b.total - a.total);
}

function soma_(items, tipo) {
  return items
    .filter(item => item.Tipo === tipo)
    .reduce((sum, item) => sum + number_(item.Valor), 0);
}

function agruparDespesasPorCategoria_(items) {
  const map = {};
  items
    .filter(item => item.Tipo === 'Despesa')
    .forEach(item => {
      const categoria = item.Categoria || 'Sem categoria';
      map[categoria] = (map[categoria] || 0) + number_(item.Valor);
    });
  return Object.keys(map)
    .map(categoria => ({ categoria, total: map[categoria] }))
    .sort((a, b) => b.total - a.total);
}

function number_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  const text = String(value).replace(/[R$\s]/g, '');
  const hasComma = text.indexOf(',') >= 0;
  const hasDot = text.indexOf('.') >= 0;
  let normalized = text;

  if (hasComma && hasDot) {
    normalized = text.replace(/\./g, '').replace(',', '.');
  } else if (hasComma) {
    normalized = text.replace(',', '.');
  }

  const parsed = Number(normalized);
  return isNaN(parsed) ? 0 : parsed;
}

function unique_(values) {
  return values.filter((value, index, array) => value && array.indexOf(value) === index);
}

function normalizarSaidaLancamento_(item) {
  const data = parseDate_(item.Data);
  if (!item.ID || !data || isNaN(data.getTime())) return null;

  return {
    id: item.ID,
    data: Utilities.formatDate(data, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    tipo: item.Tipo,
    categoria: item.Categoria,
    subcategoria: item.Subcategoria,
    descricao: item.Descricao,
    forma: item.Forma,
    conta: item.Conta,
    valor: number_(item.Valor),
    mes: isMesKey_(item.Mes) ? item.Mes : mesKey_(data),
    status: item.Status,
    observacoes: item.Observacoes
  };
}

function formatCurrency_(value) {
  return 'R$ ' + number_(value).toFixed(2).replace('.', ',');
}
