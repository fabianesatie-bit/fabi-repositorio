function buscarRascunhosApuracoes() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!emailLogado) return [];

    const usuarioLogadoObj = obterUsuarioLogado(emailLogado);
    if (usuarioLogadoObj.role === 'BLOQUEADO' || usuarioLogadoObj.role === 'GERENTE_LOJA') return [];

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Rascunhos_Apuracoes');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rascunhos = [];
    for (let i = 1; i < data.length; i++) {
      const emailCriador = data[i][31] ? data[i][31].toString().toLowerCase().trim() : '';
      const regCaso = data[i][5] ? normalizarTexto(data[i][5]) : '';
      const dirCaso = data[i][4] ? normalizarTexto(data[i][4]) : '';

      let podeVer = false;
      if (usuarioLogadoObj.role === 'MASTER' || usuarioLogadoObj.role === 'COMPLIANCE') {
        podeVer = true;
      } else if (usuarioLogadoObj.role === 'LIDERANCA') {
        if (usuarioLogadoObj.regionais.includes(regCaso) || usuarioLogadoObj.diretoria.includes(dirCaso)) podeVer = true;
      } else {
        if (emailCriador === emailLogado || usuarioLogadoObj.cargo.includes('coord')) podeVer = true;
      }

      if (!podeVer) continue;

      let dataReg = data[i][1];
      if (dataReg instanceof Date) dataReg = dataReg.toLocaleDateString('pt-BR');
      let dataRec = data[i][11];
      if (dataRec instanceof Date) dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      let dataFin = data[i][12];
      if (dataFin instanceof Date) dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');

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
    let sheet = ss.getSheetByName('Rascunhos_Apuracoes') || ss.insertSheet('Rascunhos_Apuracoes');
    
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
          rowIndex = i + 1;
          linkDocExistente = vals[i][8] ? vals[i][8].toString() : '';
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      idRascunho = 'RASC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const nomesDenunciadosArray = dados.denunciados.map(d => d.nome).filter(n => n);
    const stringDenunciadosUnificada = nomesDenunciadosArray.join(', ') || 'Não Informado';

    let linkDoc = linkDocExistente;
    let statusEstado = (tipoAcao === 'comite') ? 'Sob Análise Comitê' : 'Rascunho Pessoal';

    if (tipoAcao === 'comite') {
      linkDoc = criarOuAtualizarDocPreliminarComite(idRascunho, linkDocExistente, dados, filial, diretoria, regional, stringDenunciadosUnificada, contatos);
    }
    
    const rowData = [
      idRascunho, dataAtual, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', statusEstado,
      linkDoc, dados.nota_humor_antes || '', dados.tratativa || '', dados.data_recebimento || '', dados.data_finalizacao || '', dados.resumo || '',
      dados.diagnostico || '', dados.justificativa || '', dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente || '',
      dados.gerente_id || '', dados.gerente_nome || '', dados.agendar_intervencao || 'nao', dados.detalhes_intervencao || '',
      dados.denunciados[0]?.id || '', dados.denunciados[0]?.nome || '', dados.denunciados[0]?.filial || '',
      dados.denunciados[1]?.id || '', dados.denunciados[1]?.nome || '', dados.denunciados[1]?.filial || '',
      dados.denunciados[2]?.id || '', dados.denunciados[2]?.nome || '', dados.denunciados[2]?.filial || '',
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

    if (linkExistente && linkExistente.includes('document/d/')) {
      const match = linkExistente.match(/[-\w]{25,}/);
      if (match) docId = match[0];
      try { if (docId) doc = DocumentApp.openById(docId); } catch(e){}
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
    return '';
  }
}

function dispararAlertaComitePreliminar(contatos, filial, status, linkDoc, origem, denunciadosStr) {
  const lista = [];
  if (origem === 'Interna') {
    if (contatos.coordenador) contatos.coordenador.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.regionalEmail) contatos.regionalEmail.split(',').forEach(e => lista.push(e.trim()));
  } else {
    if (contatos.coordenador) contatos.coordenador.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.gerenteGP) contatos.gerenteGP.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.regionalEmail) contatos.regionalEmail.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.compliance) contatos.compliance.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.diretorRH) contatos.diretorRH.split(',').forEach(e => lista.push(e.trim()));
  }

  const destinatarios = [...new Set(lista)].filter(e => e);
  if (destinatarios.length === 0) return;

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

    let diretoria = 'Não identificada', regional = 'Não identificada';
    let emailCoord = '', emailGerGP = '', emailRegGP = '', emailGerenteLoja = '', emailCompliance = '', emailDirRH = '', emailDirOp = '';

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
    const anexoTexto = [dados.anexo || '', linksEvidenciasDrive].filter(x => x).join(' \n');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let apuracaoSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
    
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
    const nomesDenunciadosArray = dados.denunciados.map(d => d.nome).filter(n => n);
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
          break;
        }
      }
    }

    let docLink = docLinkExistente;

    if (docLinkExistente && docLinkExistente.includes('document/d/')) {
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
        } catch(errDoc){}
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

      aplicarPermissoesArquivo(newDocId, emailCoord, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailGerGP, 'LEITOR');
      if (dados.origem === 'Canal') {
        aplicarPermissoesArquivo(newDocId, emailCompliance, 'LEITOR');
        aplicarPermissoesArquivo(newDocId, emailDirRH, 'LEITOR');
      }
      if (emailRegGP) aplicarPermissoesArquivo(newDocId, emailRegGP, 'LEITOR');
    }

    if (!idApuracaoDefinitiva || rowIndex === -1) {
      idApuracaoDefinitiva = 'APU-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idApuracaoDefinitiva, dataRegistro, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', dados.conclusao,
      docLink, dados.nota_humor_antes || '', dados.tratativa, dados.data_recebimento, dados.data_finalizacao, dados.resumo,
      dados.diagnostico, dados.justificativa, dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente,
      dados.gerente_id, dados.gerente_nome, dados.agendar_intervencao || 'nao', dados.detalhes_intervencao,
      dados.denunciados[0]?.id || '', dados.denunciados[0]?.nome || '', dados.denunciados[0]?.filial || '',
      dados.denunciados[1]?.id || '', dados.denunciados[1]?.nome || '', dados.denunciados[1]?.filial || '',
      dados.denunciados[2]?.id || '', dados.denunciados[2]?.nome || '', dados.denunciados[2]?.filial || '',
      new Date().toLocaleString('pt-BR'), emailLogado
    ];

    if (rowIndex !== -1) {
      apuracaoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      apuracaoSheet.appendRow(rowData);
    }

    if (dados.idRascunho) {
      let rascSheet = ss.getSheetByName('Rascunhos_Apuracoes');
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
      let fbSheet = ss.getSheetByName('Feedbacks_Gerentes') || ss.insertSheet('Feedbacks_Gerentes');
      if (fbSheet.getLastRow() === 0) {
        fbSheet.appendRow(['ID Feedback', 'Filial', 'Data Envio', 'Feedback Solicitado', 'E-mail Gerente', 'Status', 'Data Resposta', 'Considerações Gerente', 'Links Anexos', 'Nome Gerente', 'ID Gerente']);
      }
      
      const feedbackId = 'FB-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      fbSheet.appendRow([feedbackId, filial, new Date(), dados.feedback_gerente, emailGerenteLoja || 'Não cadastrado', 'Pendente', '', '', '', dados.gerente_nome || '', dados.gerente_id || '']);

      const webAppUrl = ScriptApp.getService().getUrl();
      const feedbackLink = webAppUrl + "?page=gerente&idFeedback=" + feedbackId;
      
      const copiados = [Session.getActiveUser().getEmail()];
      if (emailRegGP) emailRegGP.split(',').forEach(e => { if (e.trim()) copiados.push(e.trim()) });
      const copiasCC = [...new Set(copiados)].join(',');

      if (emailGerenteLoja) enviarEmailGerenteFeedback(emailGerenteLoja, feedbackId, dados.feedback_gerente, feedbackLink, filial, copiasCC);
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
    const usuarioLogadoObj = obterUsuarioLogado(emailLogado);
    
    if (usuarioLogadoObj.role === 'BLOQUEADO' || usuarioLogadoObj.role === 'GERENTE_LOJA') {
      return [];
    }

    const lista = [];

    const dataAp = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataAp && dataAp.length > 1) {
      for (let i = 1; i < dataAp.length; i++) {
        const emailCriador = dataAp[i][32] ? dataAp[i][32].toString().toLowerCase().trim() : '';
        const dirCaso = dataAp[i][4] ? normalizarTexto(dataAp[i][4]) : '';
        const regCaso = dataAp[i][5] ? normalizarTexto(dataAp[i][5]) : '';

        let podeVer = false;
        if (usuarioLogadoObj.role === 'MASTER' || usuarioLogadoObj.role === 'COMPLIANCE') {
          podeVer = true;
        } else if (usuarioLogadoObj.role === 'LIDERANCA') {
          if (usuarioLogadoObj.regionais.includes(regCaso) || usuarioLogadoObj.diretoria.includes(dirCaso)) podeVer = true;
        } else {
          if (emailCriador === emailLogado || usuarioLogadoObj.cargo.includes('coord')) podeVer = true;
        }

        if (podeVer) {
          let dt = dataAp[i][1];
          if (dt instanceof Date) dt = dt.toLocaleDateString('pt-BR');
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
        
        let podeVer = false;
        if (usuarioLogadoObj.role === 'MASTER' || usuarioLogadoObj.role === 'COMPLIANCE') podeVer = true;
        else if (usuarioLogadoObj.role === 'LIDERANCA') podeVer = true; // Liderança ve clima geral
        else if (emailCriador === emailLogado || usuarioLogadoObj.cargo.includes('coord')) podeVer = true;

        if (podeVer) {
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
        // Para a tabela normal de GP, mostramos tudo que ele criou
        if (usuarioLogadoObj.role === 'MASTER' || usuarioLogadoObj.role === 'COMPLIANCE' || emailCriador === emailLogado || usuarioLogadoObj.cargo.includes('coord')) {
          let dt = dataDes[i][1];
          if (dt instanceof Date) dt = dt.toLocaleDateString('pt-BR');
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
    if (!id) return { erro: 'ID do registro não fornecido para edição.' };
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const usuarioLogadoObj = obterUsuarioLogado(emailLogado);
    const idBuscado = id.toString().trim().toUpperCase();

    const isReadOnlyMode = (usuarioLogadoObj.role === 'COMPLIANCE' || usuarioLogadoObj.role === 'LIDERANCA');

    const dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      for (let i = 1; i < dataApuracao.length; i++) {
        if (dataApuracao[i][0] && dataApuracao[i][0].toString().trim().toUpperCase() === idBuscado) {
          const emailCriador = dataApuracao[i][32] ? dataApuracao[i][32].toString().toLowerCase().trim() : '';
          
          if (usuarioLogadoObj.role !== 'MASTER' && usuarioLogadoObj.role !== 'COMPLIANCE' && usuarioLogadoObj.role !== 'LIDERANCA') {
            if (emailCriador && emailCriador !== emailLogado && !usuarioLogadoObj.cargo.includes('coord')) {
              return { erro: 'Acesso Negado: Você não possui privilégio de edição para este registro.' };
            }
          }

          let dataRec = dataApuracao[i][11];
          if (dataRec instanceof Date) dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          
          let dataFin = dataApuracao[i][12];
          if (dataFin instanceof Date) dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');

          return {
            tipo: 'apuracao',
            somenteLeitura: isReadOnlyMode,
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

    const dataDesligamento = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_DESLIGAMENTO_F');
    if (dataDesligamento && dataDesligamento.length > 1) {
      for (let i = 1; i < dataDesligamento.length; i++) {
        if (dataDesligamento[i][0] && dataDesligamento[i][0].toString().trim().toUpperCase() === idBuscado) {
          return {
            tipo: 'desligamento',
            somenteLeitura: isReadOnlyMode,
            id: dataDesligamento[i][0].toString(),
            filial: dataDesligamento[i][2] ? dataDesligamento[i][2].toString() : '',
            colaboradorNome: dataDesligamento[i][3] ? dataDesligamento[i][3].toString() : '',
            colaboradorId: dataDesligamento[i][4] ? dataDesligamento[i][4].toString() : '',
            tempoEmpresa: dataDesligamento[i][5] ? dataDesligamento[i][5].toString() : '',
            colaboradorCargo: dataDesligamento[i][6] ? dataDesligamento[i][6].toString() : '',
            resultados: dataDesligamento[i][7] ? dataDesligamento[i][7].toString() : '',
            justificativa: dataDesligamento[i][8] ? dataDesligamento[i][8].toString() : '',
            evidencias: dataDesligamento[i][9] ? dataDesligamento[i][9].toString() : '',
            parecerCoordenador: dataDesligamento[i][10] ? dataDesligamento[i][10].toString() : '',
            coordenadorNome: dataDesligamento[i][18] ? dataDesligamento[i][18].toString() : '',
            linkDoc: dataDesligamento[i][20] ? dataDesligamento[i][20].toString() : ''
          };
        }
      }
    }

    const dataIntervencao = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataIntervencao && dataIntervencao.length > 1) {
      for (let i = 1; i < dataIntervencao.length; i++) {
        if (dataIntervencao[i][0] && dataIntervencao[i][0].toString().trim().toUpperCase() === idBuscado) {
          let dataProg = dataIntervencao[i][6];
          if (dataProg instanceof Date) dataProg = Utilities.formatDate(dataProg, Session.getScriptTimeZone(), 'yyyy-MM-dd');

          return {
            tipo: 'intervencao',
            somenteLeitura: isReadOnlyMode,
            id: dataIntervencao[i][0].toString(),
            filial: dataIntervencao[i][1] ? dataIntervencao[i][1].toString() : '',
            linkDoc: dataIntervencao[i][2] ? dataIntervencao[i][2].toString() : '',
            statusEvolucao: dataIntervencao[i][3] ? dataIntervencao[i][3].toString() : '',
            detalhes: dataIntervencao[i][4] ? dataIntervencao[i][4].toString() : '',
            notaHumorDepois: dataIntervencao[i][5] ? dataIntervencao[i][5].toString() : '',
            novaDataProgramada: dataProg ? dataProg.toString() : '',
            colaboradores: [
              { id: dataIntervencao[i][7] ? dataIntervencao[i][7].toString() : '', nome: dataIntervencao[i][8] ? dataIntervencao[i][8].toString() : '', filial: dataIntervencao[i][9] ? dataIntervencao[i][9].toString() : '' },
              { id: dataIntervencao[i][10] ? dataIntervencao[i][10].toString() : '', nome: dataIntervencao[i][11] ? dataIntervencao[i][11].toString() : '', filial: dataIntervencao[i][12] ? dataIntervencao[i][12].toString() : '' },
              { id: dataIntervencao[i][13] ? dataIntervencao[i][13].toString() : '', nome: dataIntervencao[i][14] ? dataIntervencao[i][14].toString() : '', filial: dataIntervencao[i][15] ? dataIntervencao[i][15].toString() : '' }
            ]
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
    const userObj = obterUsuarioLogado(emailLogado);
    const idBuscado = payload.id.toString().trim().toUpperCase();
    const motivo = payload.motivo || 'Lançado Incorretamente / Anulado';

    if (userObj.role === 'COMPLIANCE' || userObj.role === 'LIDERANCA' || userObj.role === 'GERENTE_LOJA') {
       return { sucesso: false, erro: 'Acesso Negado. O seu perfil possui visualização restrita ou de leitura apenas.' };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName('HISTORICO_APURACAO');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          sheet.getRange(i + 1, 8).setValue('CANCELADO');
          sheet.getRange(i + 1, 16).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    // Mesma logica para Desligamento
    sheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          sheet.getRange(i + 1, 13).setValue('CANCELADO');
          sheet.getRange(i + 1, 16).setValue('CANCELADO');
          sheet.getRange(i + 1, 9).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    sheet = ss.getSheetByName('Intervencoes_Feedback');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          sheet.getRange(i + 1, 4).setValue('CANCELADO');
          sheet.getRange(i + 1, 5).setValue('CANCELADO / MOTIVO: ' + motivo);
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
          let apurSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
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
