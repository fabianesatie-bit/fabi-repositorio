/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Write.gs
 * Subpasta Monorepo: src/portal-gp360/
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
      return { sucesso: false, mensagem: 'Usuário sem permissão de gravação no Portal GP 360.' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!aba) {
      return { sucesso: false, mensagem: 'Aba DADOS_LANCAMENTOS não localizada na planilha master.' };
    }

    var evidenciaUrl = dados.evidenciaUrlDirect || '';
    if (fileData && fileData.base64) {
      var folder = DriveApp.getFolderById(EVIDENCIAS_FOLDER_ID);
      var bytes = Utilities.base64Decode(fileData.base64.split(',')[1]);
      var blob = Utilities.newBlob(bytes, fileData.mimeType, 'evidencia_' + new Date().getTime() + '.' + fileData.ext);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      evidenciaUrl = file.getUrl();
    }

    var novoId = 'ACT-' + new Date().getTime();
    var dataHora = new Date();

    var rate = parseFloat(dados.rateKm) || 1.20;
    var km = parseFloat(dados.kmPercorrido) || 0;
    var custoKm = km * rate;

    var alimentacao = parseFloat(dados.valorAlimentacao) || 0;
    var hospedagem = parseFloat(dados.valorHospedagem) || 0;
    var aereo = parseFloat(dados.valorAereo) || 0;
    var pedagio = parseFloat(dados.valorPedagio) || 0;
    var estacionamento = parseFloat(dados.valorEstacionamento) || 0;

    var custoTotal = custoKm + alimentacao + hospedagem + aereo + pedagio + estacionamento;

    var moedasGeradas = 1;
    var abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
    if (abaPremios) {
      var dP = abaPremios.getDataRange().getValues();
      var motUpper = String(dados.motivo || '').trim().toUpperCase();
      for (var p = 1; p < dP.length; p++) {
        if (String(dP[p][0]).trim().toUpperCase() === motUpper) {
          moedasGeradas = parseFloat(dP[p][1]) || 1;
          break;
        }
      }
    }

    var motivoLower = String(dados.motivo || '').toLowerCase();
    var ehEspecialista = motivoLower.includes('atendimento social') || 
                         motivoLower.includes('apuraç') || 
                         motivoLower.includes('apurac') || 
                         motivoLower.includes('feedback') || 
                         motivoLower.includes('acompanhamento');

    // Define o status inicial da Coluna 28 (AB)
    var statusValidacaoInicial = ehEspecialista ? 'PENDENTE' : 'VALIDADO';

    var linha28Colunas = [
      novoId,                                              // 1. ID_Lancamento (Col A)
      dataHora,                                            // 2. Data_Hora (Col B)
      controle.email,                                      // 3. Coordenador_Email (Col C)
      dados.filial || 'REGIONAL',                          // 4. Destino_Filial_Regional (Col D)
      dados.motivo || 'Atividade Operacional',             // 5. Motivo_Meta (Col E)
      dados.valorAntes || '',                              // 6. Valor_Antes (Col F)
      dados.valorDepois || '',                             // 7. Valor_Depois (Col G)
      dados.checklistNota || '',                           // 8. Checklist_Nota_Media (Col H)
      custoTotal,                                          // 9. Total_Gastos (Col I)
      moedasGeradas,                                       // 10. Moedas_Geradas (Col J)
      dados.observacao || '',                              // 11. Observacoes (Col K)
      evidenciaUrl,                                        // 12. Link_Evidencia (Col L)
      dados.dataAtividade || dataHora,                     // 13. Data_Ini (Col M)
      dados.dataAtividade || dataHora,                     // 14. Data_Fim (Col N)
      rate,                                                // 15. Valor_Km (Col O)
      km,                                                  // 16. Qde_Km (Col P)
      custoKm,                                             // 17. Total Km (Col Q)
      dados.tipoRoteiro || 'Presencial (Visita in loco)',  // 18. Tipo_Roteiro (Col R)
      dados.tema || '',                                    // 19. Sub temas (Col S)
      dados.pessoasImpactadas || 0,                        // 20. Pessoas Impactas (Col T)
      dados.tempoGasto || 0,                               // 21. Tempo Gasto (Col U)
      custoTotal,                                          // 22. Total despesa (Col V)
      alimentacao,                                         // 23. Alimentação (Col W)
      hospedagem,                                          // 24. Hospedagem (Col X)
      aereo,                                               // 25. Aereo (Col Y)
      pedagio,                                             // 26. Pedagio (Col Z)
      estacionamento,                                      // 27. Estacionamento (Col AA)
      statusValidacaoInicial                               // 28. Status_Validacao_Especialista (Col AB)
    ];

    aba.appendRow(linha28Colunas);

    if (dados.filial) {
      CacheService.getScriptCache().remove('IND_LOJA_' + dados.filial);
    }
    registrarAuditoria('REGISTRO_ATIVIDADE', 'ID: ' + novoId + ' - Loja/Destino: ' + dados.filial + ' - Status: ' + statusValidacaoInicial);

    return { 
      sucesso: true, 
      mensagem: 'Atividade registrada com sucesso!', 
      id: novoId,
      statusValidacao: statusValidacaoInicial 
    };
  } catch (err) {
    return { sucesso: false, mensagem: 'Erro na gravação: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function excluirLancamento(idLancamento) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { sucesso: false, mensagem: 'Servidor ocupado.' };
  }

  try {
    var controle = obterControleAcesso();
    if (!controle.temAcesso) return { sucesso: false, mensagem: 'Acesso negado.' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var aba = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!aba) return { sucesso: false, mensagem: 'Aba não encontrada.' };

    var dados = aba.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idLancamento)) {
        var autor = String(dados[i][2]).toLowerCase().trim();
        if (autor === controle.email.toLowerCase() || controle.isSuperAdmin) {
          aba.deleteRow(i + 1);
          registrarAuditoria('EXCLUSAO_ATIVIDADE', 'ID deletado: ' + idLancamento);
          return { sucesso: true, mensagem: 'Lançamento excluído com sucesso.' };
        } else {
          return { sucesso: false, mensagem: 'Permissão insuficiente para excluir este registro.' };
        }
      }
    }
    return { sucesso: false, mensagem: 'Registro não localizado no banco de dados.' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao excluir: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}
