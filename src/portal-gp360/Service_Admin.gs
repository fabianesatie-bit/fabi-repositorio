/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Admin.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

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

function publicarAviso(mensagem, diretoriaAlvo) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('DADOS_AVISOS');
  if (!aba) aba = ss.insertSheet('DADOS_AVISOS');

  aba.appendRow([new Date(), mensagem, controle.nome, diretoriaAlvo || 'TODAS']);
  return { sucesso: true, mensagem: 'Aviso publicado no mural!' };
}
