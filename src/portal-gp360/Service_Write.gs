/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Write.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Registra atividade no DADOS_LANCAMENTOS com trava de concorrência e upload de mídia
 */
function registrarAtividade(dados, fileData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { sucesso: false, mensagem: 'Servidor ocupado. Tente novamente em alguns segundos.' };
  }

  try {
    var controle = obterControleAcesso();
    if (!controle.temAcesso) {
      return { sucesso: false, mensagem: 'Usuário sem permissão de gravação.' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('DADOS_LANCAMENTOS');

    var evidenciaUrl = dados.evidenciaUrlDirect || '';
    if (fileData && fileData.base64) {
      var folder = DriveApp.getFolderById(EVIDENCIAS_FOLDER_ID);
      var bytes = Utilities.base64Decode(fileData.base64.split(',')[1]);
      var blob = Utilities.newBlob(bytes, fileData.mimeType, 'evidencia_' + new Date().getTime() + '.' + fileData.ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      evidenciaUrl = file.getUrl();
    }

    var novoId = Utilities.getUuid();
    var dataHora = new Date();

    // Cálculo de Custos
    var km = parseFloat(dados.kmPercorrido) || 0;
    var pedagio = parseFloat(dados.valorPedagio) || 0;
    var alimentacao = parseFloat(dados.valorAlimentacao) || 0;
    var hospedagem = parseFloat(dados.valorHospedagem) || 0;
    var outros = parseFloat(dados.outrosCustos) || 0;
    var custoTotal = (km * 1.10) + pedagio + alimentacao + hospedagem + outros; // Taxa KM Padrão

    aba.appendRow([
      novoId,
      dataHora,
      controle.email,
      dados.filial,
      dados.lojaNome || '',
      dados.motivo,
      dados.tema || '',
      dados.observacao || '',
      evidenciaUrl,
      km,
      pedagio,
      alimentacao,
      hospedagem,
      outros,
      custoTotal,
      controle.nome,
      controle.cargo,
      controle.regionais.join(', ')
    ]);

    CacheService.getScriptCache().remove('IND_LOJA_' + dados.filial);
    registrarAuditoria('REGISTRO_ATIVIDADE', 'ID: ' + novoId + ' - Loja: ' + dados.filial);

    return { sucesso: true, mensagem: 'Atividade e reembolso registrados com sucesso!', id: novoId };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro na gravação: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exclui um lançamento do histórico
 */
function excluirLancamento(idLancamento) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { sucesso: false, mensagem: 'Servidor ocupado.' };
  }

  try {
    var controle = obterControleAcesso();
    if (!controle.temAcesso) return { sucesso: false, mensagem: 'Não autorizado.' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('DADOS_LANCAMENTOS');
    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idLancamento)) {
        var autor = String(dados[i][2]).toLowerCase();
        if (autor === controle.email.toLowerCase() || controle.isSuperAdmin) {
          aba.deleteRow(i + 1);
          registrarAuditoria('EXCLUSAO_ATIVIDADE', 'ID deletado: ' + idLancamento);
          return { sucesso: true, mensagem: 'Lançamento excluído com sucesso.' };
        } else {
          return { sucesso: false, mensagem: 'Você não tem permissão para excluir este registro.' };
        }
      }
    }
    return { sucesso: false, mensagem: 'Registro não localizado.' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao excluir: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}
