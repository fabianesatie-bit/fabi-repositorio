/**
 * ============================================================================
 * DASHBOARD E CONTROLE DE BANCO DE HORAS - GOOGLE APPS SCRIPT
 * ============================================================================
 */

// 1. MENU PERSONALIZADO NA PLANILHA
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Banco de Horas')
    .addItem('🔄 Atualizar Dashboard', 'atualizarDashboard')
    .addItem('📸 Capturar Snapshot Diário Agora', 'executarSnapshotDiario')
    .addSeparator()
    .addItem('⏰ Configurar Trigger Automático (10:30h)', 'configurarTriggerDiario')
    .addToUi();
}

// 2. CAPTURA DE SNAPSHOT DIÁRIO (RODA ÀS 10:30H)
function executarSnapshotDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extracaoSheet = ss.getSheetByName('Extração 1');
  
  if (!extracaoSheet) {
    SpreadsheetApp.getUi().alert('Aba "Extração 1" não foi encontrada!');
    return;
  }

  const lastRow = extracaoSheet.getLastRow();
  if (lastRow < 2) return;

  // Leitura em Lote dos 3.900 registros
  const data = extracaoSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const dataHoje = new Date().toISOString().split('T')[0];

  // Agregação por Diretoria e Regional
  const mapa = {};

  data.forEach(row => {
    const diretoria = String(row[2] || '').trim().toUpperCase();
    const regional = String(row[3] || '').trim().toUpperCase();
    const horasFechamento = parseNum(row[7]);
    const atual = parseNum(row[8]);
    const ajustePendente = parseNum(row[9]);

    const chave = `${diretoria}||${regional}`;

    if (!mapa[chave]) {
      mapa[chave] = {
        data: dataHoje,
        diretoria: diretoria,
        regional: regional,
        totalColaboradores: 0,
        horasFechamento: 0,
        horasAtual: 0,
        ajustePendente: 0
      };
    }

    mapa[chave].totalColaboradores += 1;
    mapa[chave].horasFechamento += horasFechamento;
    mapa[chave].horasAtual += atual;
    mapa[chave].ajustePendente += ajustePendente;
  });

  // Gravar na aba Historico_Diario
  let histSheet = ss.getSheetByName('Historico_Diario');
  if (!histSheet) {
    histSheet = ss.insertSheet('Historico_Diario');
    histSheet.getRange('A1:H1').setValues([[
      'Data_Snapshot', 'Diretoria', 'Regional', 'Total_Colaboradores',
      'Horas_Fechamento', 'Horas_Atual', 'Ajuste_Pendente', 'Variacao_Delta'
    ]]).setFontWeight('bold').setBackground('#E0E0E0');
  }

  const rows = Object.values(mapa).map(item => [
    item.data,
    item.diretoria,
    item.regional,
    item.totalColaboradores,
    item.horasFechamento,
    item.horasAtual,
    item.ajustePendente,
    item.horasAtual - item.horasFechamento
  ]);

  const lastHistRow = histSheet.getLastRow();
  histSheet.getRange(lastHistRow + 1, 1, rows.length, 8).setValues(rows);

  atualizarDashboard();
}

// 3. RENDERIZAÇÃO DO DASHBOARD NA PLANILHA
function atualizarDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const extracaoSheet = ss.getSheetByName('Extração 1');
  let dashSheet = ss.getSheetByName('Dashboard_Banco_Horas');

  if (!dashSheet) {
    dashSheet = ss.insertSheet('Dashboard_Banco_Horas', 0);
  }

  // Leitura dos Filtros
  const filtroDiretoria = String(dashSheet.getRange('B2').getValue() || 'TODAS').toUpperCase().trim();
  const filtroRegional = String(dashSheet.getRange('B3').getValue() || 'TODAS').toUpperCase().trim();

  const lastRow = extracaoSheet.getLastRow();
  if (lastRow < 2) return;

  const data = extracaoSheet.getRange(2, 1, lastRow - 1, 10).getValues();

  let saldoAtualTotal = 0;
  let saldoFechamentoTotal = 0;
  let ajustesPendentesTotal = 0;
  let totalColaboradores = 0;

  const mapaRegional = {};

  data.forEach(row => {
    const diretoria = String(row[2] || '').trim().toUpperCase();
    const regional = String(row[3] || '').trim().toUpperCase();

    const bateDiretoria = filtroDiretoria === 'TODAS' || diretoria === filtroDiretoria;
    const bateRegional = filtroRegional === 'TODAS' || regional === filtroRegional;

    if (bateDiretoria && bateRegional) {
      const atual = parseNum(row[8]);
      const fechamento = parseNum(row[7]);
      const pendente = parseNum(row[9]);

      saldoAtualTotal += atual;
      saldoFechamentoTotal += fechamento;
      ajustesPendentesTotal += pendente;
      totalColaboradores += 1;
    }

    if (bateDiretoria) {
      mapaRegional[regional] = (mapaRegional[regional] || 0) + parseNum(row[8]);
    }
  });

  const variacaoAbsoluta = saldoAtualTotal - saldoFechamentoTotal;
  const variacaoPercentual = saldoFechamentoTotal > 0 ? (variacaoAbsoluta / saldoFechamentoTotal) * 100 : 0;

  // Layout do Dashboard
  dashSheet.getRange('A1:G1').merge().setValue('📊 DASHBOARD DE BANCO DE HORAS').setFontSize(16).setFontWeight('bold').setBackground('#1F497D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  dashSheet.getRange('A2').setValue('Diretoria:').setFontWeight('bold');
  dashSheet.getRange('A3').setValue('Regional:').setFontWeight('bold');

  if (!dashSheet.getRange('B2').getValue()) dashSheet.getRange('B2').setValue('TODAS');
  if (!dashSheet.getRange('B3').getValue()) dashSheet.getRange('B3').setValue('TODAS');

  // Cards de KPIs
  dashSheet.getRange('A5:B5').merge().setValue('SALDO ATUAL TOTAL').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  dashSheet.getRange('A6:B6').merge().setValue(saldoAtualTotal).setNumberFormat('#,##0.00 "hrs"').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  dashSheet.getRange('C5:D5').merge().setValue('EVOLUÇÃO vs FECHAMENTO').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  const sinal = variacaoAbsoluta >= 0 ? '+' : '';
  const corEvolucao = variacaoAbsoluta > 0 ? '#C6EFCE' : '#FFC7CE';
  dashSheet.getRange('C6:D6').merge().setValue(`${sinal}${variacaoAbsoluta.toFixed(2)} hrs (${sinal}${variacaoPercentual.toFixed(1)}%)`).setFontSize(12).setFontWeight('bold').setBackground(corEvolucao).setHorizontalAlignment('center');

  dashSheet.getRange('E5:F5').merge().setValue('AJUSTES PENDENTES (FALTAS)').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  dashSheet.getRange('E6:F6').merge().setValue(ajustesPendentesTotal).setNumberFormat('#,##0.00 "hrs"').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  dashSheet.getRange('G5').setValue('COLABORADORES').setFontWeight('bold').setBackground('#D9E1F2').setHorizontalAlignment('center');
  dashSheet.getRange('G6').setValue(totalColaboradores).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');

  // Tabela de Apoio para Gráfico
  const ranking = Object.entries(mapaRegional).sort((a, b) => b[1] - a[1]);
  dashSheet.getRange('A9:B9').setValues([['Regional', 'Saldo Atual (hrs)']]).setFontWeight('bold').setBackground('#E0E0E0');
  dashSheet.getRange('A10:B50').clearContent();

  if (ranking.length > 0) {
    dashSheet.getRange(10, 1, ranking.length, 2).setValues(ranking);
  }

  // Atualizar Gráficos Nativos
  atualizarGraficosNativos(dashSheet, ranking.length);

  ss.toast('Dashboard atualizado!', 'Sucesso', 3);
}

function atualizarGraficosNativos(sheet, totalRegionais) {
  const charts = sheet.getCharts();
  charts.forEach(c => sheet.removeChart(c));

  if (totalRegionais > 0) {
    const rangeRegional = sheet.getRange(`A9:B${Math.min(9 + totalRegionais, 25)}`);
    const chartRegional = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(rangeRegional)
      .setPosition(10, 4, 0, 0)
      .setOption('title', 'Ranking de Banco de Horas por Regional')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#1F497D'])
      .build();

    sheet.insertChart(chartRegional);
  }
}

function parseNum(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const num = parseFloat(String(val).replace(',', '.'));
  return isNaN(num) ? 0 : num;
}

// 4. CONFIGURAÇÃO DO TRIGGER AUTOMÁTICO
function configurarTriggerDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'executarSnapshotDiario') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('executarSnapshotDiario')
    .timeBased()
    .atHour(10)
    .nearMinute(30)
    .everyDays(1)
    .create();

  SpreadsheetApp.getUi().alert('Trigger diário configurado para rodar automaticamente às 10:30h!');
}
