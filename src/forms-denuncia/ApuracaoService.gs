// =============================================================================
// REGRAS DE NEGÓCIO DE APURAÇÃO, RASCUNHOS DE COMITÊ E EDIÇÃO IN-PLACE
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: ApuracaoService.gs
// =============================================================================

function buscarRascunhosApuracoes() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!emailLogado) {
      return [];
    }

    const ehMaster = verificarEhAdminMaster(emailLogado);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Rascunhos_Apuracoes');
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }
    
    const rascunhos = [];
    for (let i = 1; i < data.length; i++) {
      const emailCriador = data[i][31] ? data[i][31].toString().toLowerCase().trim() : '';

      // RESTRIÇÃO RIGOROSA DE PROPRIEDADE
      if (!ehMaster && emailCriador !== emailLogado) {
        continue;
      }

      let dataReg = data[i][1];
      if (dataReg instanceof Date) {
        dataReg = dataReg.toLocaleDateString('pt-BR');
      }
      
      let dataRec = data[i][11];
      if (dataRec instanceof Date) {
        dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      
      let dataFin = data[i][12];
      if (dataFin instanceof Date) {
        dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }

      rascunhos.push({
        idRascunho: data[i][0] ? data[i][0].toString() : '',
        dataRegistro: dataReg ? dataReg.toString() : '',
        origem: data[i][2] ? data[i][2].toString() : 'Canal',
        filial: data[i][3] ? data[i][3].toString() : '',
        apurador: data[i][6] ? data[i][6].toString() : 'GP',
        conclusao: data[i][7] ? data[i][7].toString() : 'Rascunho Pessoal',
        linkDoc: data[i][8] ? data[i][8].toString() : '',
        notaHumorAntes: data[i][9] ? data[i][9].toString() : '',
        dataRecebimento: dataRec ? dataRec.toString() : '',
        dataFinalizacao: dataFin ? dataFin.toString() : '',
        resumo: data[i][13] ? data[i][13].toString() : '',
        diagnostico: data[i][14] ? data[i][14].toString() : '',
        justificativa: data[i][15] ? data[i][15].toString() : '',
        tratativa: data[i][10] ? data[i][10].toString() : '',
        enviarFeedbackGerente: data[i][16] ? data[i][16].toString() : 'nao',
        feedbackGerenteText: data[i][17] ? data[i][17].toString() : '',
        gerenteId: data[i][18] ? data[i][18].toString() : '',
        gerenteNome: data[i][19] ? data[i][19].toString() : '',
        agendarIntervencao: data[i][20] ? data[i][20].toString() : 'nao',
        detalhesIntervencao: data[i][21] ? data[i][21].toString() : '',
        denunciados: [
          { id: data[i][22] ? data[i][22].toString() : '', nome: data[i][23] ? data[i][23].toString() : '', filial: data[i][24] ? data[i][24].toString() : '' },
          { id: data[i][25] ? data[i][25].toString() : '', nome: data[i][26] ? data[i][26].toString() : '', filial: data[i][27] ? data[i][27].toString() : '' },
          { id: data[i][28] ? data[i][28].toString() : '', nome: data[i][29] ? data[i][29].toString() : '', filial: data[i][30] ? data[i][30].toString() : '' }
        ],
        emailCriador: emailCriador
      });
    }
    return rascunhos;
  } catch (e) {
    Logger.log("Erro ao buscar rascunhos: " + e.message);
    return [];
  }
}

function salvarRascunhoApuracao(dados, tipoAcao) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Rascunhos_Apuracoes') || ss.insertSheet('Rascunhos_Apuracoes');
    
    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);
    const diretoria = contatos ? contatos.diretoria : '';
    const regional = contatos ? contatos.regional : '';
    const dataAtual = new Date().toLocaleDateString('pt-BR');
    
    let idRascunho = dados.idRascunho;
    let rowIndex = -1;
    let linkDocExistente = '';
    
    if (idRascunho) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0].toString() === idRascunho.toString()) {
          const criadorOriginal = vals[i][31] ? vals[i][31].toString().toLowerCase().trim() : '';
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode alterar este rascunho.' };
          }
          rowIndex = i + 1;
          linkDocExistente = vals[i][8] ? vals[i][8].toString() : '';
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      idRascunho = 'RASC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const nomesDenunciadosArray = dados.denunciados ? dados.denunciados.map(function(d) { return d.nome; }).filter(function(n) { return n; }) : [];
    const stringDenunciadosUnificada = nomesDenunciadosArray.join(', ') || 'Não Informado';

    let linkDoc = linkDocExistente;
    const statusEstado = (tipoAcao === 'comite') ? 'Sob Análise Comitê' : 'Rascunho Pessoal';

    if (tipoAcao === 'comite') {
      linkDoc = criarOuAtualizarDocPreliminarComite(idRascunho, linkDocExistente, dados, filial, diretoria, regional, stringDenunciadosUnificada, contatos);
    }
    
    const rowData = [
      idRascunho, dataAtual, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', statusEstado,
      linkDoc, dados.nota_humor_antes || '', dados.tratativa || '', dados.data_recebimento || '', dados.data_finalizacao || '', dados.resumo || '',
      dados.diagnostico || '', dados.justificativa || '', dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente || '',
      dados.gerente_id || '', dados.gerente_nome || '', dados.agendar_intervencao || 'nao', dados.detalhes_intervencao || '',
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].id : '',
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].nome : '',
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].filial : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].id : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].nome : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].filial : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].id : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].nome : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].filial : '',
      emailLogado
    ];
    
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    if (tipoAcao === 'comite' && contatos) {
      dispararAlertaComitePreliminar(contatos, filial, statusEstado, linkDoc, dados.origem, stringDenunciadosUnificada);
    }
    
    return { sucesso: true, idRascunho: idRascunho, linkDoc: linkDoc, tipoAcao: tipoAcao };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function criarOuAtualizarDocPreliminarComite(idRascunho, linkExistente, dados, filial, diretoria, regional, denunciadosStr, contatos) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    let doc = null;
    let docId = '';

    if (linkExistente && linkExistente.indexOf('document/d/') !== -1) {
      const match = linkExistente.match(/[-\w]{25,}/);
      if (match) {
        docId = match[0];
      }
      try {
        if (docId) {
          doc = DocumentApp.openById(docId);
        }
      } catch (e) {}
    }

    if (!doc) {
      const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
      const newFileName = "[PRELIMINAR - COMITÊ] Relatorio F." + filial + " - " + denunciadosStr + " - " + new Date().toLocaleDateString('pt-BR');
      const copiedFile = templateFile.makeCopy(newFileName, folder);
      docId = copiedFile.getId();
      doc = DocumentApp.openById(docId);
    }

    const body = doc.getBody();
    body.setText(''); 

    body.appendParagraph("REGISTRO PRELIMINAR - REVISÃO DE COMITÊ").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Atenção: Este documento é um RASCUNHO EM ANÁLISE e não possui validade de encerramento definitivo.\n");

    body.appendParagraph("1. DADOS DE CONTROLE").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Filial: " + filial + " | Diretoria: " + diretoria + " | Regional: " + regional);
    body.appendParagraph("Apurador: " + (dados.apurador || 'GP'));
    body.appendParagraph("Origem: " + (dados.origem || 'Canal'));
    body.appendParagraph("Denunciado(s): " + denunciadosStr);
    body.appendParagraph("Data Recebimento: " + (dados.data_recebimento || 'N/A') + " | Previsão Conclusão: " + (dados.data_finalizacao || 'N/A'));

    body.appendParagraph("\n2. RESUMO DA DENÚNCIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(dados.resumo || 'Em preenchimento...');

    body.appendParagraph("\n3. DIAGNÓSTICO PRELIMINAR APURADO").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(dados.diagnostico || 'Em preenchimento...');

    body.appendParagraph("\n4. PARECER E JUSTIFICATIVA TEMPORÁRIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Parecer Proposto: " + (dados.conclusao || 'Sob Análise Comitê'));
    body.appendParagraph("Justificativa: " + (dados.justificativa || 'Em preenchimento...'));
    body.appendParagraph("Plano de Ação/Tratativa Proposta: " + (dados.tratativa || 'Em preenchimento...'));

    doc.saveAndClose();

    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    aplicarPermissoesArquivo(docId, emailLogado, 'EDITOR');

    if (contatos) {
      aplicarPermissoesArquivo(docId, contatos.coordenador, 'LEITOR');
      aplicarPermissoesArquivo(docId, contatos.regionalEmail, 'LEITOR');
      if (dados.origem !== 'Interna') {
        aplicarPermissoesArquivo(docId, contatos.gerenteGP, 'LEITOR');
        aplicarPermissoesArquivo(docId, contatos.compliance, 'LEITOR');
        aplicarPermissoesArquivo(docId, contatos.diretorRH, 'LEITOR');
      }
    }

    return DriveApp.getFileById(docId).getUrl();
  } catch (e) {
    Logger.log("Erro ao gerar doc preliminar: " + e.message);
    return '';
  }
}

function dispararAlertaComitePreliminar(contatos, filial, status, linkDoc, origem, denunciadosStr) {
  const lista = [];
  if (origem === 'Interna') {
    if (contatos.coordenador) {
      contatos.coordenador.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
    if (contatos.regionalEmail) {
      contatos.regionalEmail.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
  } else {
    if (contatos.coordenador) {
      contatos.coordenador.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
    if (contatos.gerenteGP) {
      contatos.gerenteGP.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
    if (contatos.regionalEmail) {
      contatos.regionalEmail.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
    if (contatos.compliance) {
      contatos.compliance.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
    if (contatos.diretorRH) {
      contatos.diretorRH.split(',').forEach(function(e) { lista.push(e.trim()); });
    }
  }

  const destinatarios = lista.filter(function(e, i, self) {
    return e && self.indexOf(e) === i;
  });
  
  if (destinatarios.length === 0) {
    return;
  }

  const assunto = "[ANÁLISE COMITÊ] Apuração Pendente F." + filial + " (" + denunciadosStr + ")";
  const htmlBody = 
    '<div style="font-family: Arial, sans-serif; color: #1a1a1a; max-width: 600px; margin: 0 auto; border: 1px solid #c3dafe; border-radius: 8px; overflow: hidden;">' +
      '<div style="background-color: #4C51BF; padding: 20px; color: white; text-align: center;">' +
        '<div style="background-color: white; color: #4C51BF; height: 40px; width: 40px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; margin-bottom: 10px;">📋</div>' +
        '<h2 style="margin: 0; font-size: 18px; text-transform: uppercase;">Apreciação de Comitê</h2>' +
        '<p style="margin: 4px 0 0 0; font-size: 12px; font-weight: bold; opacity: 0.9;">GP & GOVERNANÇA</p>' +
      '</div><div style="padding: 24px;">' +
        '<p style="font-size: 15px;">Um relatório de apuração da <strong>Filial ' + filial + '</strong> foi enviado para avaliação e parecer do Comitê de Ética / GP.</p>' +
        '<div style="background-color: #EBF8FF; padding: 18px; border-radius: 8px; border-left: 4px solid #3182CE; margin: 20px 0;">' +
          '<p style="margin:0 0 8px 0;"><strong>Envolvido(s):</strong> ' + denunciadosStr + '</p>' +
          '<p style="margin:0;"><strong>Origem:</strong> ' + origem + '</p>' +
        '</div>' +
        '<div style="text-align: center; margin: 30px 0;"><a href="' + linkDoc + '" target="_blank" rel="noopener noreferrer" style="background-color: #4C51BF; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold;">Visualizar Rascunho do Comitê</a></div>' +
      '</div></div>';

  MailApp.sendEmail({ to: destinatarios.join(','), subject: assunto, htmlBody: htmlBody });
}

function processarNovaApuracao(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    
    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);

    let diretoria = 'Não identificada';
    let regional = 'Não identificada';
    let emailCoord = '';
    let emailGerGP = '';
    let emailRegGP = '';
    let emailGerenteLoja = '';
    let emailCompliance = '';
    let emailDirRH = '';
    let emailDirOp = '';

    if (contatos) {
      diretoria = contatos.diretoria || diretoria;
      regional = contatos.regional || regional;
      emailCoord = contatos.coordenador;
      emailGerGP = contatos.gerenteGP;
      emailRegGP = contatos.regionalEmail;
      emailGerenteLoja = contatos.gerenteEmail || contatos.gerenteLoja;
      emailCompliance = contatos.compliance;
      emailDirRH = contatos.diretorRH;
      emailDirOp = contatos.diretorOp;
    }

    let linksEvidenciasDrive = '';
    if (dados.arquivosEvidencias && dados.arquivosEvidencias.length > 0) {
      linksEvidenciasDrive = salvarEvidenciasDrive(dados.arquivosEvidencias, filial);
    }
    const anexoTexto = [dados.anexo || '', linksEvidenciasDrive].filter(function(x) { return x; }).join(' \n');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const apuracaoSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
    
    if (apuracaoSheet.getLastRow() === 0) {
      apuracaoSheet.appendRow([
        'ID Apuracao', 'Data Registro', 'Origem', 'Filial', 'Diretoria', 'Regional', 'Apurador', 'Conclusao', 'Link Doc', 'Nota Humor Antes', 
        'Tratativa', 'Data Recebimento', 'Data Conclusao', 'Resumo Denuncia', 'Diagnostico Apurado', 'Justificativa', 'Enviar Feedback Gerente', 
        'Feedback Texto', 'ID Gerente', 'Nome Gerente', 'Agendar Intervencao', 'Detalhes Intervencao', 
        'ID Denunciado 1', 'Nome Denunciado 1', 'Filial Denunciado 1',
        'ID Denunciado 2', 'Nome Denunciado 2', 'Filial Denunciado 2',
        'ID Denunciado 3', 'Nome Denunciado 3', 'Filial Denunciado 3',
        'Data Ultima Alteracao', 'Email Criador'
      ]);
    }

    const dataRegistro = new Date().toLocaleDateString('pt-BR');
    const nomesDenunciadosArray = dados.denunciados ? dados.denunciados.map(function(d) { return d.nome; }).filter(function(n) { return n; }) : [];
    const stringDenunciadosUnificada = nomesDenunciadosArray.join(', ') || 'Não Informado';

    let idApuracaoDefinitiva = dados.idApuracao || dados.idRascunho || '';
    let rowIndex = -1;
    let docLinkExistente = '';

    if (idApuracaoDefinitiva) {
      const apData = apuracaoSheet.getDataRange().getValues();
      for (let i = 1; i < apData.length; i++) {
        if (apData[i][0] && apData[i][0].toString().trim().toUpperCase() === idApuracaoDefinitiva.trim().toUpperCase()) {
          rowIndex = i + 1;
          docLinkExistente = apData[i][8] ? apData[i][8].toString() : '';
          const criadorOriginal = apData[i][32] ? apData[i][32].toString().toLowerCase().trim() : '';
          
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode editar este registro.' };
          }
          break;
        }
      }
    }

    let docLink = docLinkExistente;

    if (docLinkExistente && docLinkExistente.indexOf('document/d/') !== -1) {
      const match = docLinkExistente.match(/[-\w]{25,}/);
      if (match) {
        try {
          const doc = DocumentApp.openById(match[0]);
          const body = doc.getBody();
          body.setText(''); 

          body.appendParagraph("RELATÓRIO DE APURAÇÃO E DOSSIÊ FINAL").setHeading(DocumentApp.ParagraphHeading.HEADING1);
          body.appendParagraph("1. DADOS DE CONTROLE").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph("Filial: " + filial + " | Diretoria: " + diretoria + " | Regional: " + regional);
          body.appendParagraph("Apurador: " + (dados.apurador || 'GP'));
          body.appendParagraph("Origem: " + (dados.origem || 'Canal'));
          body.appendParagraph("Denunciado(s): " + stringDenunciadosUnificada);
          body.appendParagraph("Data Recebimento: " + (dados.data_recebimento || 'N/A') + " | Data Conclusão: " + (dados.data_finalizacao || 'N/A'));

          body.appendParagraph("\n2. RESUMO DA DENÚNCIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph(dados.resumo || '');

          body.appendParagraph("\n3. DIAGNÓSTICO APURADO").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph(dados.diagnostico || '');

          body.appendParagraph("\n4. PARECER FINAL E JUSTIFICATIVA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph("Parecer Conclusivo: " + (dados.conclusao || ''));
          body.appendParagraph("Justificativa: " + (dados.justificativa || ''));
          body.appendParagraph("Plano de Ação/Tratativa: " + (dados.tratativa || ''));
          body.appendParagraph("Evidências / Anexos: " + (anexoTexto || 'Sem evidências adicionais no Drive.'));

          doc.saveAndClose();
          aplicarPermissoesArquivo(match[0], emailLogado, 'EDITOR');
        } catch(errDoc) {}
      }
    } else {
      const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const newFileName = "Relatorio de Apuração F." + filial + " - " + stringDenunciadosUnificada + " - " + dataRegistro;
      const copiedFile = templateFile.makeCopy(newFileName, folder);
      const newDocId = copiedFile.getId();

      const doc = DocumentApp.openById(newDocId);
      const body = doc.getBody();

      body.replaceText('{{filial}}', filial);
      body.replaceText('{{diretoria}}', diretoria);
      body.replaceText('{{regional}}', regional);
      body.replaceText('{{denunciados}}', stringDenunciadosUnificada);
      body.replaceText('{{apurador}}', dados.apurador || '');
      body.replaceText('{{data_recebimento}}', dados.data_recebimento || '');
      body.replaceText('{{data_finalizacao}}', dados.data_finalizacao || '');
      body.replaceText('{{resumo}}', dados.resumo || '');
      body.replaceText('{{diagnostico}}', dados.diagnostico || '');
      body.replaceText('{{conclusao}}', dados.conclusao || '');
      body.replaceText('{{justificativa}}', dados.justificativa || '');
      body.replaceText('{{tratativa}}', dados.tratativa || '');
      body.replaceText('{{gerente_id}}', dados.gerente_id || '');
      body.replaceText('{{gerente_nome}}', dados.gerente_nome || '');
      body.replaceText('{{anexo}}', anexoTexto || 'Sem evidências adicionais no Drive.');
      doc.saveAndClose();
      
      docLink = copiedFile.getUrl();

      aplicarPermissoesArquivo(newDocId, emailLogado, 'EDITOR');
      aplicarPermissoesArquivo(newDocId, emailCoord, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailGerGP, 'LEITOR');
      if (dados.origem === 'Canal') {
        aplicarPermissoesArquivo(newDocId, emailCompliance, 'LEITOR');
        aplicarPermissoesArquivo(newDocId, emailDirRH, 'LEITOR');
      }
      if (emailRegGP) {
        aplicarPermissoesArquivo(newDocId, emailRegGP, 'LEITOR');
      }
    }

    if (!idApuracaoDefinitiva || rowIndex === -1) {
      idApuracaoDefinitiva = 'APU-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idApuracaoDefinitiva, dataRegistro, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', dados.conclusao,
      docLink, dados.nota_humor_antes || '', dados.tratativa, dados.data_recebimento, dados.data_finalizacao, dados.resumo,
      dados.diagnostico, dados.justificativa, dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente,
      dados.gerente_id, dados.gerente_nome, dados.agendar_intervencao || 'nao', dados.detalhes_intervencao,
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].id : '',
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].nome : '',
      dados.denunciados && dados.denunciados[0] ? dados.denunciados[0].filial : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].id : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].nome : '',
      dados.denunciados && dados.denunciados[1] ? dados.denunciados[1].filial : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].id : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].nome : '',
      dados.denunciados && dados.denunciados[2] ? dados.denunciados[2].filial : '',
      new Date().toLocaleString('pt-BR'), emailLogado
    ];

    if (rowIndex !== -1) {
      apuracaoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      apuracaoSheet.appendRow(rowData);
    }

    if (dados.idRascunho) {
      const rascSheet = ss.getSheetByName('Rascunhos_Apuracoes');
      if (rascSheet) {
        const rascVals = rascSheet.getDataRange().getValues();
        for (let i = 1; i < rascVals.length; i++) {
          if (rascVals[i][0].toString().trim().toUpperCase() === dados.idRascunho.toString().trim().toUpperCase()) {
            rascSheet.deleteRow(i + 1);
            break;
          }
        }
      }
    }

    if (dados.enviar_feedback_gerente === 'sim') {
      const fbSheet = ss.getSheetByName('Feedbacks_Gerentes') || ss.insertSheet('Feedbacks_Gerentes');
      if (fbSheet.getLastRow() === 0) {
        fbSheet.appendRow(['ID Feedback', 'Filial', 'Data Envio', 'Feedback Solicitado', 'E-mail Gerente', 'Status', 'Data Resposta', 'Considerações Gerente', 'Links Anexos', 'Nome Gerente', 'ID Gerente']);
      }
      
      const feedbackId = 'FB-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      fbSheet.appendRow([feedbackId, filial, new Date(), dados.feedback_gerente, emailGerenteLoja || 'Não cadastrado', 'Pendente', '', '', '', dados.gerente_nome || '', dados.gerente_id || '']);

      const webAppUrl = ScriptApp.getService().getUrl();
      const feedbackLink = webAppUrl + "?page=gerente&idFeedback=" + feedbackId;
      
      const copiados = [Session.getActiveUser().getEmail()];
      if (emailRegGP) {
        emailRegGP.split(',').forEach(function(e) { if (e.trim()) copiados.push(e.trim()); });
      }
      const copiasCC = copiados.filter(function(e, i, self) { return e && self.indexOf(e) === i; }).join(',');

      if (emailGerenteLoja) {
        enviarEmailGerenteFeedback(emailGerenteLoja, feedbackId, dados.feedback_gerente, feedbackLink, filial, copiasCC);
      }
    }

    return { sucesso: true, link: docLink, id: idApuracaoDefinitiva };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}

function listarTodosRegistrosUsuario() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const lista = [];

    const dataAp = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataAp && dataAp.length > 1) {
      for (let i = 1; i < dataAp.length; i++) {
        const emailCriador = dataAp[i][32] ? dataAp[i][32].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          let dt = dataAp[i][1];
          if (dt instanceof Date) {
            dt = dt.toLocaleDateString('pt-BR');
          }
          lista.push({
            tipo: 'apuracao',
            tipoLabel: 'Apuração',
            id: dataAp[i][0] ? dataAp[i][0].toString() : '',
            data: dt ? dt.toString() : 'N/A',
            filial: dataAp[i][3] ? dataAp[i][3].toString() : '',
            status: dataAp[i][7] ? dataAp[i][7].toString() : 'Concluído',
            resumo: dataAp[i][13] ? dataAp[i][13].toString() : 'N/A',
            emailCriador: emailCriador
          });
        }
      }
    }

    const dataInt = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataInt && dataInt.length > 1) {
      for (let i = 1; i < dataInt.length; i++) {
        const emailCriador = dataInt[i][16] ? dataInt[i][16].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          lista.push({
            tipo: 'intervencao',
            tipoLabel: 'Feedback Clima',
            id: dataInt[i][0] ? dataInt[i][0].toString() : '',
            data: 'N/A',
            filial: dataInt[i][1] ? dataInt[i][1].toString() : '',
            status: dataInt[i][3] ? dataInt[i][3].toString() : 'Registrado',
            resumo: dataInt[i][4] ? dataInt[i][4].toString() : 'N/A',
            emailCriador: emailCriador
          });
        }
      }
    }

    const dataDes = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_DESLIGAMENTO_F');
    if (dataDes && dataDes.length > 1) {
      for (let i = 1; i < dataDes.length; i++) {
        const emailCriador = dataDes[i][21] ? dataDes[i][21].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          let dt = dataDes[i][1];
          if (dt instanceof Date) {
            dt = dt.toLocaleDateString('pt-BR');
          }
          lista.push({
            tipo: 'desligamento',
            tipoLabel: 'Desligamento',
            id: dataDes[i][0] ? dataDes[i][0].toString() : '',
            data: dt ? dt.toString() : 'N/A',
            filial: dataDes[i][2] ? dataDes[i][2].toString() : '',
            status: dataDes[i][12] ? dataDes[i][12].toString() : 'Pendente',
            resumo: "Colaborador: " + (dataDes[i][3] ? dataDes[i][3].toString() : 'N/A'),
            emailCriador: emailCriador
          });
        }
      }
    }

    return lista;
  } catch (e) {
    return [];
  }
}

function buscarRegistroParaEdicao(id) {
  try {
    if (!id) {
      return { erro: 'ID do registro não fornecido para edição.' };
    }
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const idBuscado = id.toString().trim().toUpperCase();

    const dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      for (let i = 1; i < dataApuracao.length; i++) {
        if (dataApuracao[i][0] && dataApuracao[i][0].toString().trim().toUpperCase() === idBuscado) {
          const emailCriador = dataApuracao[i][32] ? dataApuracao[i][32].toString().toLowerCase().trim() : '';
          
          if (!ehMaster && emailCriador && emailCriador !== emailLogado) {
            return { erro: 'Acesso Negado: Apenas o criador original deste registro pode editá-lo.' };
          }

          let dataRec = dataApuracao[i][11];
          if (dataRec instanceof Date) {
            dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }
          
          let dataFin = dataApuracao[i][12];
          if (dataFin instanceof Date) {
            dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          }

          return {
            tipo: 'apuracao',
            id: dataApuracao[i][0].toString(),
            dataRegistro: dataApuracao[i][1] ? dataApuracao[i][1].toString() : '',
            origem: dataApuracao[i][2] ? dataApuracao[i][2].toString() : 'Canal',
            filial: dataApuracao[i][3] ? dataApuracao[i][3].toString() : '',
            diretoria: dataApuracao[i][4] ? dataApuracao[i][4].toString() : '',
            regional: dataApuracao[i][5] ? dataApuracao[i][5].toString() : '',
            apurador: dataApuracao[i][6] ? dataApuracao[i][6].toString() : '',
            conclusao: dataApuracao[i][7] ? dataApuracao[i][7].toString() : '',
            linkDoc: dataApuracao[i][8] ? dataApuracao[i][8].toString() : '',
            notaHumorAntes: dataApuracao[i][9] ? dataApuracao[i][9].toString() : '',
            tratativa: dataApuracao[i][10] ? dataApuracao[i][10].toString() : '',
            dataRecebimento: dataRec ? dataRec.toString() : '',
            dataFinalizacao: dataFin ? dataFin.toString() : '',
            resumo: dataApuracao[i][13] ? dataApuracao[i][13].toString() : '',
            diagnostico: dataApuracao[i][14] ? dataApuracao[i][14].toString() : '',
            justificativa: dataApuracao[i][15] ? dataApuracao[i][15].toString() : '',
            enviarFeedbackGerente: dataApuracao[i][16] ? dataApuracao[i][16].toString() : 'nao',
            feedbackGerenteText: dataApuracao[i][17] ? dataApuracao[i][17].toString() : '',
            gerenteId: dataApuracao[i][18] ? dataApuracao[i][18].toString() : '',
            gerenteNome: dataApuracao[i][19] ? dataApuracao[i][19].toString() : '',
            agendarIntervencao: dataApuracao[i][20] ? dataApuracao[i][20].toString() : 'nao',
            detalhesIntervencao: dataApuracao[i][21] ? dataApuracao[i][21].toString() : '',
            denunciados: [
              { id: dataApuracao[i][22] ? dataApuracao[i][22].toString() : '', nome: dataApuracao[i][23] ? dataApuracao[i][23].toString() : '', filial: dataApuracao[i][24] ? dataApuracao[i][24].toString() : '' },
              { id: dataApuracao[i][25] ? dataApuracao[i][25].toString() : '', nome: dataApuracao[i][26] ? dataApuracao[i][26].toString() : '', filial: dataApuracao[i][27] ? dataApuracao[i][27].toString() : '' },
              { id: dataApuracao[i][28] ? dataApuracao[i][28].toString() : '', nome: dataApuracao[i][29] ? dataApuracao[i][29].toString() : '', filial: dataApuracao[i][30] ? dataApuracao[i][30].toString() : '' }
            ],
            emailCriador: emailCriador
          };
        }
      }
    }

    return { erro: 'Registro ' + id + ' não localizado nas bases.' };
  } catch (e) {
    return { erro: 'Erro ao buscar registro: ' + e.toString() };
  }
}

function cancelarRegistroProcesso(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const idBuscado = payload.id.toString().trim().toUpperCase();
    const motivo = payload.motivo || 'Lançado Incorretamente / Anulado';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName('HISTORICO_APURACAO');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][32] ? vals[i][32].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este registro.' };
          }
          sheet.getRange(i + 1, 8).setValue('CANCELADO');
          sheet.getRange(i + 1, 16).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    sheet = ss.getSheetByName('Rascunhos_Apuracoes');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][31] ? vals[i][31].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este rascunho.' };
          }

          const apurSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
          apurSheet.appendRow([
            vals[i][0], vals[i][1], vals[i][2], vals[i][3], vals[i][4], vals[i][5], vals[i][6], 'CANCELADO',
            vals[i][8], vals[i][9], vals[i][10], vals[i][11], vals[i][12], vals[i][13],
            vals[i][14], 'CANCELADO / MOTIVO: ' + motivo, vals[i][16], vals[i][17],
            vals[i][18], vals[i][19], vals[i][20], vals[i][21],
            vals[i][22], vals[i][23], vals[i][24],
            vals[i][25], vals[i][26], vals[i][27],
            vals[i][28], vals[i][29], vals[i][30],
            new Date().toLocaleString('pt-BR'), criador
          ]);

          sheet.deleteRow(i + 1);
          return { sucesso: true };
        }
      }
    }

    return { sucesso: false, erro: 'Registro não localizado para anulação.' };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function estruturarTextoNativo(textoBruto) {
  if (!textoBruto) {
    return {
      resumo: '',
      diagnostico: '',
      justificativa: '',
      tratativa: '',
      feedback_gerente: ''
    };
  }

  const texto = textoBruto.toString().trim();
  let resumo = '';
  let diagnostico = '';
  let justificativa = '';
  let tratativa = '';
  let feedback_gerente = '';

  if (texto.indexOf('#') !== -1) {
    const blocos = texto.split('#');
    blocos.forEach(function(b) {
      if (!b.trim()) {
        return;
      }
      const primeiraEspaco = b.trim().indexOf(' ');
      const tag = primeiraEspaco !== -1 ? b.trim().substring(0, primeiraEspaco).toLowerCase() : b.trim().toLowerCase();
      const conteudo = primeiraEspaco !== -1 ? b.trim().substring(primeiraEspaco).trim() : '';

      if (tag.indexOf('relato') !== -1 || tag.indexOf('resumo') !== -1 || tag.indexOf('denuncia') !== -1) {
        resumo = conteudo;
      } else if (tag.indexOf('diag') !== -1 || tag.indexOf('fato') !== -1 || tag.indexOf('constata') !== -1) {
        diagnostico = conteudo;
      } else if (tag.indexOf('just') !== -1 || tag.indexOf('parecer') !== -1 || tag.indexOf('fundam') !== -1) {
        justificativa = conteudo;
      } else if (tag.indexOf('acao') !== -1 || tag.indexOf('tratativa') !== -1 || tag.indexOf('plano') !== -1) {
        tratativa = conteudo;
      } else if (tag.indexOf('feed') !== -1 || tag.indexOf('gerente') !== -1 || tag.indexOf('diretri') !== -1) {
        feedback_gerente = conteudo;
      }
    });
  }

  if (!resumo || !diagnostico) {
    const matchFilial = texto.match(/(?:loja|filial)\s*(\d+)/i);
    const numFilial = matchFilial ? matchFilial[1] : '';

    const matchDenunciante = texto.match(/denunciante\s*([a-zà-ú]+)/i) || texto.match(/colaborador\s*([a-zà-ú]+)/i);
    const nomeDenunciante = matchDenunciante ? matchDenunciante[1].toUpperCase() : 'Denunciante';

    const matchDenunciado = texto.match(/lideran[çc]a\s*([a-zà-ú]+)/i) || texto.match(/gerente\s*([a-zà-ú]+)/i);
    const nomeDenunciado = matchDenunciado ? matchDenunciado[1].toUpperCase() : 'Gestora Local';

    const temRelacionamento = /relacionamento|ficam|carro|bolsa|esposo|marido|bar/i.test(texto);

    if (temRelacionamento) {
      if (!resumo) {
        resumo = "Acolheu-se o relato do colaborador " + nomeDenunciante + (numFilial ? " (Filial " + numFilial + ")" : "") +
          ", alegando postura diferenciada, tratamento desproporcional e interferência na distribuição de atendimentos por parte da liderança direta, Sra. " + nomeDenunciado + ".";
      }

      if (!diagnostico) {
        diagnostico = "Realizadas entrevistas detalhadas com a equipe local e envolvidos. Constatou-se relacionamento interpessoal prévio e desalinhamento de conduta operacional no salão de vendas.";
      }
    } else {
      const paragrafos = texto.split('\n').map(function(p) { return p.trim(); }).filter(function(p) { return p.length > 0; });
      if (!resumo && paragrafos.length >= 1) {
        resumo = paragrafos[0];
      }
      if (!diagnostico && paragrafos.length >= 2) {
        diagnostico = paragrafos.slice(1, Math.min(4, paragrafos.length)).join('\n\n');
      }
    }
  }

  if (!resumo) {
    resumo = texto.substring(0, 400);
  }
  if (!diagnostico) {
    diagnostico = "Com base nos relatos colhidos, identificou-se desalinhamento de conduta e fragilidade na comunicação direta.";
  }
  
  if (!justificativa) {
    justificativa = "A apuração evidencia inconformidade com as diretrizes do Código de Ética e Conduta Corporativo.";
  }

  if (!tratativa) {
    tratativa = "1. Readequação de Subordinação Direta.\n2. Orientação Pedagógica à Liderança.\n3. Monitoramento Contínuo de Clima.";
  }

  if (!feedback_gerente) {
    feedback_gerente = "Prezado(a) Gestor(a),\n\nAlinhamos as diretrizes executivas para cumprimento imediato em sua unidade.";
  }

  return {
    resumo: resumo,
    diagnostico: diagnostico,
    justificativa: justificativa,
    tratativa: tratativa,
    feedback_gerente: feedback_gerente
  };
}
