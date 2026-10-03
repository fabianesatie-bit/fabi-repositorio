/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS
 * Arquivo: Service_Write.gs (Gravação no Banco de Dados com Concorrência Segura)
 */

function verificarAbaFolgas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG.nomeAbaFolgas);
  if (!aba) {
    aba = ss.insertSheet(CONFIG.nomeAbaFolgas);
    aba.appendRow(["Carimbo de Data/Hora", "Filial", "Nome", "ID do Colaborador", "Data da Folga"]);
    aba.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
  }
}

/**
 * Grava ou Atualiza as Folgas sem duplicar linhas (com Lock de Concorrência)
 */
function salvarFolgasPortal(filial, agendamentos) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Aguarda até 15s para garantir exclusividade de gravação
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    verificarAbaFolgas();
    let aba = ss.getSheetByName(CONFIG.nomeAbaFolgas);
    
    let agora = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
    let filialNorm = normalizarFilialId(filial);
    let dadosFolgas = aba.getDataRange().getValues();

    // Mapeia onde já existem registros por ID do Colaborador (Coluna D / índice 3)
    let mapaLinhasColab = {};
    for (let i = 1; i < dadosFolgas.length; i++) {
      let idCadastrado = dadosFolgas[i][3] ? String(dadosFolgas[i][3]).trim() : "";
      if (idCadastrado) {
        mapaLinhasColab[idCadastrado] = i + 1; // Linha física 1-indexed
      }
    }

    agendamentos.forEach(ag => {
      let idColab = String(ag.id).trim();
      let dataEscolhida = ag.dataFolga;

      if (mapaLinhasColab[idColab]) {
        // Se já existe agendamento anterior para esse colaborador, ATUALIZA a mesma linha
        let linha = mapaLinhasColab[idColab];
        aba.getRange(linha, 1, 1, 5).setValues([[agora, filialNorm, ag.nome, idColab, dataEscolhida]]);
      } else {
        // Se for o primeiro agendamento, insere nova linha
        aba.appendRow([agora, filialNorm, ag.nome, idColab, dataEscolhida]);
        mapaLinhasColab[idColab] = aba.getLastRow();
      }
    });

    return true;
  } catch (e) {
    Logger.log("Erro ao salvar folgas: " + e.message);
    throw e;
  } finally {
    lock.releaseLock();
  }
}
