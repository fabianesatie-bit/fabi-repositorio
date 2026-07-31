function processarNovoDesligamento(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();

    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);
    let diretoria = 'MG/CO', regional = 'Brasília', emailRegionalGP = '', emailDiretorRH = '', emailDiretorOp = '';

    if (contatos) {
      diretoria = contatos.diretoria || diretoria;
      regional = contatos.regional || regional;
      emailRegionalGP = contatos.regionalEmail;
      emailDiretorRH = contatos.diretorRH;
      emailDiretorOp = contatos.diretorOp;
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let fullDesligamentoSheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F') || ss.insertSheet('HISTORICO_DESLIGAMENTO_F');
    
    if (fullDesligamentoSheet.getLastRow() === 0) {
      fullDesligamentoSheet.appendRow([
        'ID Desligamento', 'Data Registro', 'Filial', 'Colaborador', 'ID', 'Tempo Empresa', 'Cargo', 'Resultados', 'Justificativa', 'Evidencias', 
        'Parecer Coordenador', 'Email Regional', 'Status Regional', 'Parecer Regional', 'Email Diretor Op', 'Status Diretor Op', 'Parecer Diretor Op', 
        'Email Diretor RH', 'Coordenador Nome', 'Links Imagens Resultados', 'Link Apuração', 'Email Criador'
      ]);
    }

    let idDesligamento = dados.id || '';
    let rowIndex = -1;
    let docLinkExistente = '';

    if (idDesligamento) {
      const desData = fullDesligamentoSheet.getDataRange().getValues();
      for (let i = 1; i < desData.length; i++) {
        if (desData[i][0] && desData[i][0].toString().trim().toUpperCase() === idDesligamento.trim().toUpperCase()) {
          rowIndex = i + 1;
          docLinkExistente = desData[i][20] ? desData[i][20].toString() : '';
          break;
        }
      }
    }

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const linksImagens = [];
    if (dados.arquivos_resultados && dados.arquivos_resultados.length > 0) {
      dados.arquivos_resultados.forEach((arq, index) => {
        if (arq.dados && arq.dados.includes(',')) {
          const splitData = arq.dados.split(',');
          const contentType = splitData[0].match(/:(.*?);/)[1];
          const rawData = Utilities.base64Decode(splitData[1]);
          const fileName = "Resultados_F" + filial + "_" + dados.colaborador_id + "_" + index + "_" + (arq.nome || ".png");
          linksImagens.push(folder.createFile(Utilities.newBlob(rawData, contentType, fileName)).getUrl());
        }
      });
    }

    let docLink = docLinkExistente;

    if (!docLinkExistente) {
      const templateFile = DriveApp.getFileById(TEMPLATE_DESLIGAMENTO_DOC_ID);
      const copiedFile = templateFile.makeCopy("Relatório de Desligamento F." + filial + " - " + dados.colaborador_nome, folder);
      const newDocId = copiedFile.getId();
      const doc = DocumentApp.openById(newDocId);
      const body = doc.getBody();

      body.replaceText('{{filial}}', filial);
      body.replaceText('{{diretoria}}', diretoria);
      body.replaceText('{{regional}}', regional);
      body.replaceText('{{colaborador_nome}}', dados.colaborador_nome);
      body.replaceText('{{colaborador_id}}', dados.colaborador_id); 
      body.replaceText('{{colaborador_cargo}}', dados.colaborador_cargo);
      body.replaceText('{{tempo_empresa}}', dados.tempo_empresa || 'Não informado');
      body.replaceText('{{data_registro}}', new Date().toLocaleDateString('pt-BR'));
      body.replaceText('{{coordenador_nome}}', dados.coordenador_nome);
      body.replaceText('{{resultados}}', dados.resultados);
      body.replaceText('{{resultados_imagens}}', linksImagens.length > 0 ? linksImagens.join(' \n') : 'Nenhuma imagem.');
      body.replaceText('{{justificativa}}', dados.justificativa);
      body.replaceText('{{evidencias}}', dados.evidencias || 'Sem link.');
      body.replaceText('{{parecer_coordenador}}', dados.parecer_coordenador);
      doc.saveAndClose();

      docLink = copiedFile.getUrl();

      aplicarPermissoesArquivo(newDocId, emailLogado, 'EDITOR');
      aplicarPermissoesArquivo(newDocId, emailRegionalGP, 'EDITOR'); // Editor para emitir parecer
      aplicarPermissoesArquivo(newDocId, emailDiretorOp, 'EDITOR');
      aplicarPermissoesArquivo(newDocId, emailDiretorRH, 'LEITOR');
    }

    if (!idDesligamento || rowIndex === -1) {
      idDesligamento = 'DES-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idDesligamento, new Date(), filial, dados.colaborador_nome, dados.colaborador_id, dados.tempo_empresa, dados.colaborador_cargo,
      dados.resultados, dados.justificativa, dados.evidencias || '', dados.parecer_coordenador, emailRegionalGP, 'Pendente', '',
      emailDiretorOp, 'Pendente', '', emailDiretorRH, dados.coordenador_nome, linksImagens.join(' \n'), docLink, emailLogado
    ];

    if (rowIndex !== -1) {
      fullDesligamentoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      fullDesligamentoSheet.appendRow(rowData);
      // Disparo de notificação aos Diretores/Regionais que um dossiê foi submetido.
      if (emailRegionalGP || emailDiretorOp) {
        dispararEmailAlcadaDesligamento(emailRegionalGP, emailDiretorOp, filial, dados.colaborador_nome, docLink);
      }
    }

    return { sucesso: true, link: docLink, id: idDesligamento };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function listarDesligamentosPendentesLideranca() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const userObj = obterUsuarioLogado(emailLogado);

    if (userObj.role !== 'LIDERANCA' && userObj.role !== 'MASTER') {
      return [];
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const pendentes = [];

    for (let i = 1; i < data.length; i++) {
      const emailReg = data[i][11] ? data[i][11].toString().toLowerCase().trim() : '';
      const statusReg = data[i][12] ? data[i][12].toString() : '';
      const emailDir = data[i][14] ? data[i][14].toString().toLowerCase().trim() : '';
      const statusDir = data[i][15] ? data[i][15].toString() : '';

      // Verifica se o login está nas alçadas e está pendente
      let exigeAlcadaLocal = false;
      if (emailReg === emailLogado && statusReg === 'Pendente') exigeAlcadaLocal = true;
      if (emailDir === emailLogado && statusDir === 'Pendente') exigeAlcadaLocal = true;
      
      if (userObj.role === 'MASTER' && (statusReg === 'Pendente' || statusDir === 'Pendente')) {
         exigeAlcadaLocal = true;
      }

      if (exigeAlcadaLocal) {
        let dt = data[i][1];
        if (dt instanceof Date) dt = dt.toLocaleDateString('pt-BR');
        pendentes.push({
          id: data[i][0].toString(),
          data: dt,
          filial: data[i][2].toString(),
          nome: data[i][3].toString(),
          cargo: data[i][6].toString(),
          linkDoc: data[i][20].toString()
        });
      }
    }
    return pendentes;
  } catch (e) {
    Logger.log("Erro: " + e.message);
    return [];
  }
}

function salvarParecerDesligamentoLideranca(idDesligamento, parecerTexto) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let docId = '';
    let emailCriador = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toUpperCase() === idDesligamento.toUpperCase()) {
        rowIndex = i + 1;
        const emailReg = data[i][11] ? data[i][11].toString().toLowerCase().trim() : '';
        const emailDir = data[i][14] ? data[i][14].toString().toLowerCase().trim() : '';
        emailCriador = data[i][21] ? data[i][21].toString() : '';
        const docLink = data[i][20] ? data[i][20].toString() : '';

        if (docLink && docLink.includes('document/d/')) {
          docId = docLink.match(/[-\w]{25,}/)[0];
        }

        if (emailReg === emailLogado) {
          sheet.getRange(rowIndex, 13).setValue('Aprovado/Parecer Emitido');
          sheet.getRange(rowIndex, 14).setValue(parecerTexto);
        } else if (emailDir === emailLogado) {
          sheet.getRange(rowIndex, 16).setValue('Aprovado/Parecer Emitido');
          sheet.getRange(rowIndex, 17).setValue(parecerTexto);
        } else {
          sheet.getRange(rowIndex, 16).setValue('Parecer Master');
          sheet.getRange(rowIndex, 17).setValue(parecerTexto);
        }
        break;
      }
    }

    if (rowIndex === -1) return { sucesso: false, erro: 'Registro não localizado.' };

    // Escreve o parecer diretamente no documento gerado
    if (docId) {
      try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        body.appendParagraph("\nPARECER DE LIDERANÇA / ALÇADA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph("Responsável: " + emailLogado);
        body.appendParagraph("Data: " + new Date().toLocaleString('pt-BR'));
        body.appendParagraph(parecerTexto);
        doc.saveAndClose();
      } catch (errDoc) {
        Logger.log("Falha ao registrar doc: " + errDoc.message);
      }
    }

    if (emailCriador) {
       MailApp.sendEmail({
         to: emailCriador,
         subject: "[GP] Novo Parecer de Desligamento Recebido",
         htmlBody: "A liderança acabou de inserir o parecer de desligamento no processo " + idDesligamento + ". Acesse o painel ou o relatório Docs para checar."
       });
    }

    return { sucesso: true };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function dispararEmailAlcadaDesligamento(regEmail, dirEmail, filial, nome, linkDoc) {
  const emails = [regEmail, dirEmail].filter(e => e).join(',');
  if (!emails) return;

  const assunto = "[ALÇADA PENDENTE] Relatório de Desligamento Submetido - F." + filial;
  const htmlBody = 
    '<div style="font-family: Arial; padding: 20px;">' +
    '<h2 style="color: #0086FF;">Aprovação Pendente de Desligamento</h2>' +
    '<p>Prezada Liderança, o dossiê referente ao colaborador <strong>' + nome + '</strong> da Filial <strong>' + filial + '</strong> foi gerado.</p>' +
    '<p>Acesse o painel GP Magalu para inserir seu parecer formal de aprovação.</p>' +
    '<a href="' + linkDoc + '" style="background: #0086FF; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Ler Relatório no Docs</a>' +
    '</div>';

  try { MailApp.sendEmail({ to: emails, subject: assunto, htmlBody: htmlBody }); } catch(e){}
}
