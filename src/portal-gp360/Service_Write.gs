// =============================================================================
// GRAVAÇÕES E ESCRITAS CORPORATIVAS (LOCKSERVICE BLINDADO E RÁPIDO)
// =============================================================================

function registrarAtividade(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      throw new Error("Usuário não autorizado a realizar lançamentos.");
    }
    registrarAuditoria("Nova Atividade", dados.motivo + " | " + dados.destino);
    try {
      const cache = CacheService.getScriptCache();
      if (dados.destino && !String(dados.destino).startsWith("REGIONAL") && !String(dados.destino).startsWith("DIRETORIA")) {
          let fId = parseInt(normalizarFilialId(String(dados.destino).split('-')[0].trim()), 10);
          if (!isNaN(fId)) cache.remove('IND_LOJA_' + fId);
      }
    } catch(err) {}

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DADOS_LANCAMENTOS');
    const arrayInsert = new Array(27).fill("");
    const dHora = new Date();

    arrayInsert[0] = "ACT-" + dHora.getTime();
    arrayInsert[1] = dHora;
    arrayInsert[2] = emailLogado;
    arrayInsert[3] = dados.destino;
    arrayInsert[4] = dados.motivo;
    arrayInsert[5] = "";
    arrayInsert[6] = "";
    arrayInsert[7] = ""; 
    arrayInsert[8] = "";
    arrayInsert[9] = (dados.moedas !== undefined) ? dados.moedas : 1;
    arrayInsert[10] = dados.observacoes || "";
    arrayInsert[11] = dados.linkEvidenciaExterna || "";
    arrayInsert[12] = dados.dataViagem || dados.dataAtividade || "";
    arrayInsert[13] = "";
    arrayInsert[14] = dados.kmValor || "";
    arrayInsert[15] = dados.kmQtd || "";
    arrayInsert[16] = dados.kmCusto || "";
    arrayInsert[17] = dados.roteiro || "";
    arrayInsert[18] = dados.subTema || "";
    arrayInsert[19] = dados.pessoasImpactadas || 0;
    arrayInsert[20] = dados.tempoGasto || 0;
    arrayInsert[21] = dados.gastoTotal || 0;
    arrayInsert[22] = dados.alimentacao || 0;
    arrayInsert[23] = dados.hospedagem || 0;
    arrayInsert[24] = dados.aereo || 0;
    arrayInsert[25] = dados.pedagio || 0;
    arrayInsert[26] = dados.estacionamento || 0;
    sheet.appendRow(arrayInsert);
    return { sucesso: true };
  } catch (e) { 
    return { erro: e.message }; 
  } finally {
    lock.releaseLock();
  }
}

function excluirLancamento(idLancamento) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      throw new Error("Operação não autorizada.");
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DADOS_LANCAMENTOS');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idLancamento) {
        if (String(data[i][2]).toLowerCase().trim() === emailLogado || controle.isSuperAdmin) {
            try {
                const destinoStr = String(data[i][3]).trim();
                const cache = CacheService.getScriptCache();
                if (!destinoStr.startsWith("REGIONAL") && !destinoStr.startsWith("DIRETORIA")) {
                    let fId = parseInt(normalizarFilialId(destinoStr.split('-')[0].trim()), 10);
                    if (!isNaN(fId)) cache.remove('IND_LOJA_' + fId);
                }
            } catch(e) {}
            sheet.deleteRow(i + 1); 
            return { sucesso: true };
        }
      }
    }
    return { erro: "Documento não localizado ou permissões insuficientes." };
  } catch (e) { 
    return { erro: e.message }; 
  } finally {
    lock.releaseLock();
  }
}
