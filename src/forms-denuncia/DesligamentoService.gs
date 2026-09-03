// =============================================================================
// SERVIÇO DE PROCESSAMENTO DE DESLIGAMENTOS E GOVERNANÇA DE PESSOAS
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: DesligamentoService.gs
// =============================================================================

/**
 * Processa o registro de novo Dossiê de Desligamento pelo Coordenador de GP,
 * gera o documento oficial formatado com design moderno e links clicáveis,
 * compartilha no Google Drive com os 6 stakeholders e dispara notificação por e-mail.
 */
function processarNovoDesligamento(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();

    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);
    let diretoria = 'MG/CO', regional = 'Brasília';
    let emailRegionalGP = '', emailDiretorRH = '', emailDiretorOp = '', emailGerenteGP = '', emailCoord = '', emailCompliance = '';

    if (contatos) {
      diretoria = contatos.diretoria || diretoria;
      regional = contatos.regional || regional;
      emailRegionalGP = contatos.regionalEmail || '';
      emailDiretorRH = contatos.diretorRH || '';
      emailDiretorOp = contatos.diretorOp || '';
      emailGerenteGP = contatos.gerenteGP || '';
      emailCoord = contatos.coordenador || '';
      emailCompliance = contatos.compliance || '';
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
          const fileName = "Resultados_F" + filial + "_" + dados.colaborador_id + "_" + (index + 1) + "_" + (arq.nome || ".png");
          const createdImg = folder.createFile(Utilities.newBlob(rawData, contentType, fileName));
          
          // Blindagem de privacidade
          try { createdImg.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e){}
          linksImagens.push({
            nome: arq.nome || ("Evidência " + (index + 1) + " - Print"),
            url: createdImg.getUrl()
          });
        } else if (arq.dados && arq.dados.startsWith('http')) {
          linksImagens.push({
            nome: arq.nome || ("Evidência " + (index + 1)),
            url: arq.dados
          });
        }
      });
    }

    let docLink = docLinkExistente;
    let docId = '';

    if (docLinkExistente && docLinkExistente.includes('document/d/')) {
      const match = docLinkExistente.match(/[-\w]{25,}/);
      if (match) docId = match[0];
    }

    // Geração do Documento Google Docs com Design Executivo e Moderno
    if (docId) {
      try {
        const doc = DocumentApp.openById(docId);
        construirCorpoDossieDesligamento(doc, dados, filial, diretoria, regional, linksImagens);
        doc.saveAndClose();
      } catch (errDoc) {
        Logger.log("Erro ao atualizar doc existente: " + errDoc.message);
      }
    } else {
      const nomeArquivoDoc = "Dossiê de Desligamento F." + filial + " - " + dados.colaborador_nome + " - " + new Date().toLocaleDateString('pt-BR');
      const doc = DocumentApp.create(nomeArquivoDoc);
      docId = doc.getId();
      
      // Move o arquivo recém-criado para a pasta corporativa designada
      const driveFile = DriveApp.getFileById(docId);
      folder.addFile(driveFile);
      try { DriveApp.getRootFolder().removeFile(driveFile); } catch(e){}

      construirCorpoDossieDesligamento(doc, dados, filial, diretoria, regional, linksImagens);
      doc.saveAndClose();
      docLink = driveFile.getUrl();
    }

    // =========================================================================
    // COMPARTILHAMENTO AUTOMÁTICO DO ARQUIVO COM OS 6 STAKEHOLDERS:
    // 1. Gerente de GP
    // 2. Diretor OP
    // 3. Diretor RH
    // 4. Regional
    // 5. Coordenador GP
    // 6. Compliance
    // =========================================================================
    if (contatos) {
      if (contatos.gerenteGP) aplicarPermissoesArquivo(docId, contatos.gerenteGP, 'LEITOR');
      if (contatos.diretorOp) aplicarPermissoesArquivo(docId, contatos.diretorOp, 'LEITOR');
      if (contatos.diretorRH) aplicarPermissoesArquivo(docId, contatos.diretorRH, 'LEITOR');
      if (contatos.regionalEmail) aplicarPermissoesArquivo(docId, contatos.regionalEmail, 'LEITOR');
      if (contatos.compliance) aplicarPermissoesArquivo(docId, contatos.compliance, 'LEITOR');
      if (contatos.coordenador) aplicarPermissoesArquivo(docId, contatos.coordenador, 'EDITOR');
    }
    aplicarPermissoesArquivo(docId, emailLogado, 'EDITOR');

    if (!idDesligamento || rowIndex === -1) {
      idDesligamento = 'DES-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const urlsApenasStr = linksImagens.map(l => l.url).join('\n');
    const rowData = [
      idDesligamento, new Date(), filial, dados.colaborador_nome, dados.colaborador_id, dados.tempo_empresa || 'Não informado', dados.colaborador_cargo,
      dados.resultados, dados.justificativa, dados.evidencias || '', dados.parecer_coordenador, emailRegionalGP, 'Concluído', 'Finalizado pelo Coordenador GP',
      emailDiretorOp, 'Concluído', 'Finalizado pelo Coordenador GP', emailDiretorRH, dados.coordenador_nome, urlsApenasStr, docLink, emailLogado
    ];

    if (rowIndex !== -1) {
      fullDesligamentoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      fullDesligamentoSheet.appendRow(rowData);
    }

    // =========================================================================
    // DISPARO DE E-MAIL AUTOMATIZADO COM O RELATÓRIO PARA OS 6 STAKEHOLDERS
    // =========================================================================
    dispararEmailConclusaoDesligamento(contatos, filial, diretoria, regional, dados, docLink, emailLogado);

    return { sucesso: true, link: docLink, id: idDesligamento };
  } catch (err) {
    Logger.log("Erro processarNovoDesligamento: " + err.toString());
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Constrói o corpo do Dossiê de Desligamento com design corporativo de alto padrão,
 * tipografia limpa, tabela de identificação e links 100% clicáveis.
 * Removida a coleta de pareceres do Regional e Diretor conforme solicitação.
 */
function construirCorpoDossieDesligamento(doc, dados, filial, diretoria, regional, linksImagens) {
  const body = doc.getBody();
  body.setText(''); // Limpa o corpo para reconstrução limpa e sem tags residuais

  // Configurações de margens do documento (2 cm = ~56 pt)
  body.setMarginTop(50);
  body.setMarginBottom(50);
  body.setMarginLeft(50);
  body.setMarginRight(50);

  // ---------------------------------------------------------------------------
  // CABEÇALHO CORPORATIVO
  // ---------------------------------------------------------------------------
  const pHeaderCorp = body.appendParagraph('GESTÃO DE PESSOAS | MAGAZINE LUIZA');
  pHeaderCorp.setFontFamily('Arial')
             .setFontSize(10)
             .setBold(true)
             .setForegroundColor('#160048')
             .setSpacingAfter(2);

  const pTitulo = body.appendParagraph('DOSSIÊ OFICIAL DE DESLIGAMENTO');
  pTitulo.setFontFamily('Arial')
         .setFontSize(18)
         .setBold(true)
         .setForegroundColor('#0086FF')
         .setSpacingAfter(2);

  const pSubtitulo = body.appendParagraph('Documento Oficial de Governança e Decisão de Pessoas');
  pSubtitulo.setFontFamily('Arial')
            .setFontSize(9)
            .setItalic(true)
            .setForegroundColor('#6B7280')
            .setSpacingAfter(14);

  // ---------------------------------------------------------------------------
  // SEÇÃO 1: DADOS DE IDENTIFICAÇÃO E REGISTRO (TABELA ESTRUTURADA)
  // ---------------------------------------------------------------------------
  adicionarTituloSecao(body, '1. DADOS DE IDENTIFICAÇÃO E REGISTRO');

  const tabela = body.appendTable([
    ['UNIDADE E REGIONAL', 'DADOS DO COLABORADOR', 'CONTROLE E EMISSÃO'],
    [
      'Filial: ' + filial + '\nDiretoria: ' + (diretoria || 'N/A') + '\nRegional: ' + (regional || 'N/A'),
      'Nome: ' + (dados.colaborador_nome || 'N/A') + '\nID/RE: ' + (dados.colaborador_id || 'N/A') + '\nCargo: ' + (dados.colaborador_cargo || 'N/A') + '\nTempo de Empresa: ' + (dados.tempo_empresa || 'Não informado'),
      'Data de Emissão: ' + new Date().toLocaleDateString('pt-BR') + '\nCoordenador GP: ' + (dados.coordenador_nome || 'GP') + '\nStatus: Concluído'
    ]
  ]);

  // Estilização da Tabela
  const tableAttributes = {};
  tableAttributes[DocumentApp.Attribute.BORDER_COLOR] = '#D1D5DB';
  tableAttributes[DocumentApp.Attribute.BORDER_WIDTH] = 1;
  tabela.setAttributes(tableAttributes);

  // Cabeçalho da Tabela
  const rowHeader = tabela.getRow(0);
  for (let c = 0; c < 3; c++) {
    const cell = rowHeader.getCell(c);
    cell.setBackgroundColor('#F3F4F6');
    cell.setPaddingTop(6);
    cell.setPaddingBottom(6);
    cell.setPaddingLeft(8);
    cell.setPaddingRight(8);
    const p = cell.getChild(0).asParagraph();
    p.setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#1F2937');
  }

  // Conteúdo da Tabela
  const rowData = tabela.getRow(1);
  for (let c = 0; c < 3; c++) {
    const cell = rowData.getCell(c);
    cell.setBackgroundColor('#FFFFFF');
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(8);
    cell.setPaddingRight(8);
    const p = cell.getChild(0).asParagraph();
    p.setFontFamily('Arial').setFontSize(9).setForegroundColor('#374151').setLineSpacing(1.2);
  }

  body.appendParagraph('').setSpacingAfter(8);

  // ---------------------------------------------------------------------------
  // SEÇÃO 2: RESULTADOS E HISTÓRICO DO COLABORADOR
  // ---------------------------------------------------------------------------
  adicionarTituloSecao(body, '2. HISTÓRICO DE PERFORMANCE E RESULTADOS');
  
  const pGuiaResultados = body.appendParagraph('(Histórico de metas, KPIs operacionais, conduta e entregas do profissional)');
  pGuiaResultados.setFontFamily('Arial').setFontSize(8.5).setItalic(true).setForegroundColor('#6B7280').setSpacingAfter(6);

  const pResultados = body.appendParagraph(dados.resultados || 'Sem descrição de resultados informada.');
  pResultados.setFontFamily('Arial').setFontSize(10).setForegroundColor('#1F2937').setLineSpacing(1.15).setSpacingAfter(10);

  // Bloco de Evidências Visuais / Prints com Links Clicáveis
  const pSubImg = body.appendParagraph('Evidências Visuais e Prints de Desempenho / Metas:');
  pSubImg.setFontFamily('Arial').setFontSize(9.5).setBold(true).setForegroundColor('#160048').setSpacingAfter(4);

  if (linksImagens && linksImagens.length > 0) {
    linksImagens.forEach((imgObj, idx) => {
      const pLink = body.appendParagraph('📎 ');
      pLink.setSpacingAfter(3);
      const textLabel = pLink.appendText('Evidência ' + (idx + 1) + ': ' + (imgObj.nome || 'Print de Metas / Resultados') + ' (Clique para abrir no Drive)');
      textLabel.setFontFamily('Arial')
               .setFontSize(9.5)
               .setForegroundColor('#0086FF')
               .setUnderline(true)
               .setLinkUrl(imgObj.url);
    });
  } else {
    const pSemImg = body.appendParagraph('Nenhum print ou imagem adicional anexado.');
    pSemImg.setFontFamily('Arial').setFontSize(9).setItalic(true).setForegroundColor('#9CA3AF').setSpacingAfter(6);
  }

  body.appendParagraph('').setSpacingAfter(8);

  // ---------------------------------------------------------------------------
  // SEÇÃO 3: MOTIVO E JUSTIFICATIVA DO DESLIGAMENTO
  // ---------------------------------------------------------------------------
  adicionarTituloSecao(body, '3. MOTIVO E JUSTIFICATIVA DO DESLIGAMENTO');

  const pGuiaJustificativa = body.appendParagraph('(Fundamentação técnica e contextual que baseia a tomada de decisão pelo desligamento corporativo)');
  pGuiaJustificativa.setFontFamily('Arial').setFontSize(8.5).setItalic(true).setForegroundColor('#6B7280').setSpacingAfter(6);

  const pJustificativa = body.appendParagraph(dados.justificativa || 'Sem justificativa informada.');
  pJustificativa.setFontFamily('Arial').setFontSize(10).setForegroundColor('#1F2937').setLineSpacing(1.15).setSpacingAfter(12);

  // ---------------------------------------------------------------------------
  // SEÇÃO 4: DOCUMENTOS SUPORTE E EVIDÊNCIAS (LINKS CLICÁVEIS)
  // ---------------------------------------------------------------------------
  adicionarTituloSecao(body, '4. DOCUMENTOS SUPORTE E EVIDÊNCIAS');

  const pGuiaEvidencias = body.appendParagraph('(Links do Drive contendo advertências assinadas, suspensões, feedbacks formais ou cartas de próprio punho)');
  pGuiaEvidencias.setFontFamily('Arial').setFontSize(8.5).setItalic(true).setForegroundColor('#6B7280').setSpacingAfter(6);

  formatarCampoTextoComLinksClicaveis(body, dados.evidencias);

  body.appendParagraph('').setSpacingAfter(8);

  // ---------------------------------------------------------------------------
  // SEÇÃO 5: PARECER TÉCNICO DO COORDENADOR DE GP
  // ---------------------------------------------------------------------------
  adicionarTituloSecao(body, '5. PARECER TÉCNICO DO COORDENADOR DE GP');

  const pCoordMeta = body.appendParagraph('Responsável Técnico: ' + (dados.coordenador_nome || 'Coordenador GP') + ' | Data: ' + new Date().toLocaleDateString('pt-BR'));
  pCoordMeta.setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#4B5563').setSpacingAfter(6);

  // Caixa de destaque para o parecer do coordenador
  const tabParecer = body.appendTable([[dados.parecer_coordenador || 'Parecer favorável ao desligamento registrado pelo Coordenador GP.']]);
  const pParecerAtts = {};
  pParecerAtts[DocumentApp.Attribute.BORDER_COLOR] = '#93C5FD';
  pParecerAtts[DocumentApp.Attribute.BORDER_WIDTH] = 1.5;
  tabParecer.setAttributes(pParecerAtts);

  const cellParecer = tabParecer.getRow(0).getCell(0);
  cellParecer.setBackgroundColor('#EFF6FF');
  cellParecer.setPaddingTop(10);
  cellParecer.setPaddingBottom(10);
  cellParecer.setPaddingLeft(12);
  cellParecer.setPaddingRight(12);

  const pTxtParecer = cellParecer.getChild(0).asParagraph();
  pTxtParecer.setFontFamily('Arial').setFontSize(10).setItalic(true).setForegroundColor('#1E3A8A').setLineSpacing(1.2);

  body.appendParagraph('').setSpacingAfter(14);

  // ---------------------------------------------------------------------------
  // RODAPÉ INSTITUCIONAL
  // ---------------------------------------------------------------------------
  const pDiv = body.appendParagraph('_________________________________________________________________________________');
  pDiv.setFontFamily('Arial').setFontSize(8).setForegroundColor('#E5E7EB').setSpacingAfter(6);

  const pRodape1 = body.appendParagraph('Dossiê Eletrônico de Governança de Clima e Gestão de Pessoas | Magazine Luiza S.A.');
  pRodape1.setFontFamily('Arial').setFontSize(8.5).setBold(true).setForegroundColor('#6B7280').setSpacingAfter(2);

  const pRodape2 = body.appendParagraph('Documento concluído eletronicamente via Portal GP 360º. Informações confidenciais para uso restrito da liderança autorizada.');
  pRodape2.setFontFamily('Arial').setFontSize(8).setForegroundColor('#9CA3AF').setSpacingAfter(0);
}

/**
 * Adiciona um título de seção padronizado com barra visual corporativa
 */
function adicionarTituloSecao(body, tituloTexto) {
  const p = body.appendParagraph(tituloTexto);
  p.setFontFamily('Arial')
   .setFontSize(11.5)
   .setBold(true)
   .setForegroundColor('#160048')
   .setSpacingBefore(10)
   .setSpacingAfter(4);
  return p;
}

/**
 * Converte qualquer texto com links ou URLs em parágrafos com hiperlinks clicáveis
 */
function formatarCampoTextoComLinksClicaveis(body, texto) {
  if (!texto || texto.trim() === '' || texto.trim().toLowerCase() === 'sem link.' || texto.trim().toLowerCase() === 'sem link') {
    const pVazio = body.appendParagraph('Sem links ou documentos adicionais informados.');
    pVazio.setFontFamily('Arial').setFontSize(9).setItalic(true).setForegroundColor('#9CA3AF').setSpacingAfter(6);
    return;
  }

  const linhas = texto.split('\n');
  const urlRegex = /(https?:\/\/[^\s]+)/g;

  linhas.forEach((linha, lIndex) => {
    linha = linha.trim();
    if (!linha) return;

    if (linha.match(/^https?:\/\/[^\s]+$/i)) {
      // Linha contendo apenas a URL
      const p = body.appendParagraph('🔗 ');
      p.setSpacingAfter(4);
      const linkText = p.appendText('Abrir Documento / Evidência no Drive (' + (lIndex + 1) + ')');
      linkText.setFontFamily('Arial')
               .setFontSize(9.5)
               .setForegroundColor('#0086FF')
               .setUnderline(true)
               .setLinkUrl(linha);
    } else if (urlRegex.test(linha)) {
      // Linha com texto explicativo e URL embutida
      const p = body.appendParagraph('');
      p.setSpacingAfter(4);
      p.setFontFamily('Arial').setFontSize(9.5).setForegroundColor('#1F2937');
      
      let lastIndex = 0;
      let match;
      urlRegex.lastIndex = 0;

      while ((match = urlRegex.exec(linha)) !== null) {
        if (match.index > lastIndex) {
          p.appendText(linha.substring(lastIndex, match.index));
        }
        const linkPart = p.appendText(' [Abrir Link no Drive] ');
        linkPart.setFontFamily('Arial')
                .setFontSize(9.5)
                .setForegroundColor('#0086FF')
                .setUnderline(true)
                .setLinkUrl(match[0]);
        lastIndex = match.index + match[0].length;
      }
      if (lastIndex < linha.length) {
        p.appendText(linha.substring(lastIndex));
      }
    } else {
      // Texto puro
      const p = body.appendParagraph(linha);
      p.setFontFamily('Arial').setFontSize(9.5).setForegroundColor('#374151').setSpacingAfter(4);
    }
  });
}

/**
 * Dispara e-mail corporativo com notificação e link do Dossiê para os 6 grupos:
 * - Gerente de GP
 * - Diretor OP
 * - Diretor RH
 * - Regional
 * - Coordenador GP
 * - Compliance
 */
function dispararEmailConclusaoDesligamento(contatos, filial, diretoria, regional, dados, docLink, emailLogado) {
  try {
    const listaEmails = [];
    
    if (contatos) {
      if (contatos.gerenteGP) contatos.gerenteGP.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
      if (contatos.diretorOp) contatos.diretorOp.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
      if (contatos.diretorRH) contatos.diretorRH.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
      if (contatos.regionalEmail) contatos.regionalEmail.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
      if (contatos.coordenador) contatos.coordenador.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
      if (contatos.compliance) contatos.compliance.split(',').forEach(e => listaEmails.push(e.trim().toLowerCase()));
    }
    if (emailLogado) {
      listaEmails.push(emailLogado.trim().toLowerCase());
    }

    // Remove duplicados e e-mails inválidos
    const destinatariosUnicos = [...new Set(listaEmails)].filter(e => e && e.indexOf('@') !== -1);
    if (destinatariosUnicos.length === 0) return;

    const destinatariosStr = destinatariosUnicos.join(',');
    const assunto = "[GP MAGALU] Dossiê de Desligamento Concluído - Filial " + filial + " - " + dados.colaborador_nome;

    const htmlBody = 
      '<div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; background-color: #f9fafb; border-radius: 12px; overflow: hidden; border: 1px solid #e5e7eb;">' +
        '<div style="background-color: #160048; padding: 24px; text-align: left; border-bottom: 4px solid #0086FF;">' +
          '<span style="font-size: 11px; font-weight: bold; color: #93c5fd; text-transform: uppercase; letter-spacing: 1px;">Gestão de Pessoas | Magazine Luiza</span>' +
          '<h2 style="color: #ffffff; margin: 6px 0 0 0; font-size: 20px;">Dossiê Oficial de Desligamento Concluído</h2>' +
        '</div>' +
        '<div style="padding: 24px; color: #374151; font-size: 14px; line-height: 1.6;">' +
          '<p>Olá,</p>' +
          '<p>Informamos que o <strong>Dossiê Oficial de Desligamento</strong> referente ao colaborador abaixo foi finalizado pelo Coordenador de GP responsável e já está consolidado para conhecimento e governança corporativa:</p>' +
          '<div style="background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 8px; padding: 16px; margin: 18px 0;">' +
            '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">' +
              '<tr><td style="padding: 5px 0; color: #6b7280; width: 140px;"><strong>Filial / Unidade:</strong></td><td style="padding: 5px 0; color: #111827;">Filial ' + filial + ' (' + (regional || 'Regional') + ' - ' + (diretoria || 'Diretoria') + ')</td></tr>' +
              '<tr><td style="padding: 5px 0; color: #6b7280;"><strong>Colaborador:</strong></td><td style="padding: 5px 0; color: #111827;"><strong>' + dados.colaborador_nome + '</strong> (ID: ' + dados.colaborador_id + ')</td></tr>' +
              '<tr><td style="padding: 5px 0; color: #6b7280;"><strong>Cargo:</strong></td><td style="padding: 5px 0; color: #111827;">' + dados.colaborador_cargo + '</td></tr>' +
              '<tr><td style="padding: 5px 0; color: #6b7280;"><strong>Tempo Empresa:</strong></td><td style="padding: 5px 0; color: #111827;">' + (dados.tempo_empresa || 'Não informado') + '</td></tr>' +
              '<tr><td style="padding: 5px 0; color: #6b7280;"><strong>Coordenador GP:</strong></td><td style="padding: 5px 0; color: #111827;">' + dados.coordenador_nome + ' (' + emailLogado + ')</td></tr>' +
              '<tr><td style="padding: 5px 0; color: #6b7280;"><strong>Data de Registro:</strong></td><td style="padding: 5px 0; color: #111827;">' + new Date().toLocaleDateString('pt-BR') + '</td></tr>' +
            '</table>' +
          '</div>' +
          '<div style="background-color: #eff6ff; border-left: 4px solid #0086FF; padding: 12px 16px; margin: 16px 0; border-radius: 0 8px 8px 0;">' +
            '<p style="margin: 0 0 4px 0; font-size: 12px; font-weight: bold; color: #1e3a8a; text-transform: uppercase;">Parecer do Coordenador de GP:</p>' +
            '<p style="margin: 0; font-style: italic; color: #1e40af; font-size: 13px;">"' + dados.parecer_coordenador + '"</p>' +
          '</div>' +
          '<div style="text-align: center; margin: 28px 0 16px 0;">' +
            '<a href="' + docLink + '" target="_blank" style="background-color: #0086FF; color: #ffffff; padding: 12px 28px; font-size: 14px; font-weight: bold; text-decoration: none; border-radius: 8px; display: inline-block; box-shadow: 0 2px 4px rgba(0,134,255,0.3);">' +
              '📄 Acessar Dossiê Completo no Google Docs' +
            '</a>' +
          '</div>' +
          '<p style="font-size: 11px; color: #9ca3af; text-align: center; margin-top: 20px;">' +
            'Este documento foi compartilhado automaticamente no Google Drive com os perfis autorizados (Gerente de GP, Diretor OP, Diretor RH, Regional, Coordenador GP e Compliance).' +
          '</p>' +
        '</div>' +
        '<div style="background-color: #f3f4f6; padding: 14px 24px; text-align: center; border-top: 1px solid #e5e7eb; font-size: 11px; color: #6b7280;">' +
          'Mensagem automática emitida pelo <strong>Portal GP 360º</strong> • Magazine Luiza S.A.<br>' +
          'Cultura do Jeito Luiza de Ser • Uso Corporativo e Confidencial' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: destinatariosStr,
      subject: assunto,
      htmlBody: htmlBody
    });
  } catch (e) {
    Logger.log("Erro ao disparar email de desligamento: " + e.message);
  }
}

// =============================================================================
// FUNÇÕES DE COMPATIBILIDADE (PRESERVADAS PARA EVITAR QUEBRA DE SISTEMAS EXISTENTES)
// =============================================================================

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
    Logger.log("Erro listarDesligamentosPendentesLideranca: " + e.message);
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

    if (docId) {
      try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        body.appendParagraph("\nCONSIDERAÇÕES DE LIDERANÇA").setFontFamily('Arial').setFontSize(11).setBold(true).setForegroundColor('#160048');
        body.appendParagraph("Responsável: " + emailLogado + " | Data: " + new Date().toLocaleString('pt-BR')).setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#6B7280');
        body.appendParagraph(parecerTexto).setFontFamily('Arial').setFontSize(10).setForegroundColor('#1F2937');
        doc.saveAndClose();
      } catch (errDoc) {
        Logger.log("Falha ao registrar consideração em doc: " + errDoc.message);
      }
    }

    if (emailCriador) {
       MailApp.sendEmail({
         to: emailCriador,
         subject: "[GP MAGALU] Consideração de Liderança Inserida no Dossiê",
         htmlBody: "A liderança (" + emailLogado + ") inseriu considerações no processo " + idDesligamento + ". Acesse o relatório Docs para checar."
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

  const assunto = "[GP MAGALU] Dossiê de Desligamento - Filial " + filial + " - " + nome;
  const htmlBody = 
    '<div style="font-family: Arial; padding: 20px;">' +
    '<h2 style="color: #0086FF;">Dossiê de Desligamento Registrado</h2>' +
    '<p>Prezada Liderança, o dossiê referente ao colaborador <strong>' + nome + '</strong> da Filial <strong>' + filial + '</strong> foi concluído.</p>' +
    '<a href="' + linkDoc + '" style="background: #0086FF; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">Ler Relatório no Docs</a>' +
    '</div>';

  try { MailApp.sendEmail({ to: emails, subject: assunto, htmlBody: htmlBody }); } catch(e){}
}

