/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Admin.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

function adicionarNatureza(novaNatureza) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito a administradores.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) config = ss.insertSheet('CONFIGURAÇÕES');

  config.appendRow([novaNatureza, 'NATUREZA_EXTRA', 'GERAL', novaNatureza]);
  return { sucesso: true, mensagem: 'Natureza cadastrada com sucesso!' };
}

function excluirNatureza(naturezaTexto) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso negado.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) return { sucesso: false, mensagem: 'Aba não encontrada.' };

  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var val = String(dados[i][0] || '').trim().toLowerCase();
    if (val === naturezaTexto.trim().toLowerCase()) {
      config.deleteRow(i + 1);
      return { sucesso: true, mensagem: 'Natureza removida com sucesso.' };
    }
  }
  return { sucesso: false, mensagem: 'Natureza não localizada.' };
}

function adicionarTema(categoria, temaTexto) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito a administradores.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) config = ss.insertSheet('CONFIGURAÇÕES');

  config.appendRow(['TEMA_' + categoria, 'TEMA_DROPDOWN', categoria, temaTexto]);
  return { sucesso: true, mensagem: 'Tema cadastrado com sucesso!' };
}

function excluirTema(categoria, temaTexto) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso negado.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) return { sucesso: false, mensagem: 'Aba não configurada.' };

  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var cat = String(dados[i][2] || '').trim().toUpperCase();
    var val = String(dados[i][3] || '').trim().toLowerCase();
    if (cat === categoria.toUpperCase() && val === temaTexto.trim().toLowerCase()) {
      config.deleteRow(i + 1);
      return { sucesso: true, mensagem: 'Tema removido com sucesso.' };
    }
  }
  return { sucesso: false, mensagem: 'Tema não encontrado.' };
}
