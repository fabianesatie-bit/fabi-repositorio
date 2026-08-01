/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Admin.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Adiciona uma nova natureza de atividade
 */
function adicionarNatureza(novaNatureza) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito a administradores.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) config = ss.insertSheet('CONFIGURAÇÕES');

  config.appendRow([novaNatureza, 'NATUREZA_ATIVIDADE', 'NATUREZA', novaNatureza]);
  return { sucesso: true, mensagem: 'Natureza cadastrada com sucesso!' };
}

/**
 * Exclui uma natureza de atividade
 */
function excluirNatureza(naturezaTexto) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso negado.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) return { sucesso: false, mensagem: 'Aba indisponível.' };

  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim().toLowerCase() === String(naturezaTexto).trim().toLowerCase()) {
      config.deleteRow(i + 1);
      return { sucesso: true, mensagem: 'Natureza removida.' };
    }
  }
  return { sucesso: false, mensagem: 'Item não localizado.' };
}

/**
 * Cadastra um novo Tema para Reunião Regional ou Treinamento
 */
function adicionarTema(categoria, temaTexto) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  if (!config) config = ss.insertSheet('CONFIGURAÇÕES');

  config.appendRow(['TEMA_' + categoria, 'TEMA_DROPDOWN', categoria, temaTexto]);
  return { sucesso: true, mensagem: 'Tema cadastrado com sucesso!' };
}

/**
 * Exclui um Tema cadastrado
 */
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

/**
 * Publica aviso corporativo na Home
 */
function publicarAviso(mensagem, diretoriaAlvo) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Acesso restrito.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('DADOS_AVISOS');
  if (!aba) aba = ss.insertSheet('DADOS_AVISOS');

  aba.appendRow([new Date(), mensagem, controle.nome, diretoriaAlvo || 'TODAS']);
  return { sucesso: true, mensagem: 'Aviso publicado no mural!' };
}

/**
 * Exclui aviso corporativo do mural
 */
function excluirAvisoPlanilha(dataAviso, mensagemAviso) {
  var controle = obterControleAcesso();
  if (!controle.isAdmin) return { sucesso: false, mensagem: 'Sem permissão.' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('DADOS_AVISOS');
  if (!aba) return { sucesso: false, mensagem: 'Aba de avisos indisponível.' };

  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][1]).trim() === String(mensagemAviso).trim()) {
      aba.deleteRow(i + 1);
      return { sucesso: true, mensagem: 'Aviso removido com sucesso.' };
    }
  }
  return { sucesso: false, mensagem: 'Aviso não localizado.' };
}
