// =============================================================================
// LINHA DO TEMPO DE CLIMA, INTERVENÇÕES E RESPOSTA DE GERENTES DE LOJA
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: ClimaService.gs
// =============================================================================

/**
 * MOTOR DE VISIBILIDADE PARA FILIAIS (CLIMA)
 * Garante que Coordenação e Liderança só vejam a linha do tempo de suas praças.
 */
function usuarioPodeVerFilial(userObj, regFilial, dirFilial) {
  if (userObj.role === 'MASTER' || userObj.role === 'COMPLIANCE') return true;

  const cargo = userObj.cargo ? userObj.cargo.toLowerCase() : '';
  const isGestaoLimitada = userObj.role === 'LIDERANCA' ||
                           cargo.includes('coord') ||
                           cargo.includes('gtgp') ||
                           cargo.includes('gerente gp') ||
                           cargo.includes('gerentegp');

  if (isGestaoLimitada) {
    const usrReg = normalizarTexto(userObj.regionais);
    const usrDir = normalizarTexto(userObj.diretoria);
    const fReg = normalizarTexto(regFilial);
    const fDir = normalizarTexto(dirFilial);

    if (usrReg === 'todas' || usrDir === 'todas') return true;

    // Trava de Soberania: Se possui Regionais, ignora a Diretoria
    if (usrReg !== '') {
      if (fReg !== '' && usrReg.includes(fReg)) return true;
      return false; // Bloqueia sumariamente se não bater a regional
    } 
    // Se não tem Regional, avalia pela Diretoria
    else if (usrDir !== '') {
      if (fDir !== '' && usrDir.includes(fDir)) return true;
      return false;
    }
    
    return false; // Gestão com cadastro vazio não vê nada por segurança
  }

  // Analistas GP sem cargo de liderança operam a nível Brasil para registros
  return true; 
}

/**
 * Busca a linha do tempo completa de clima para uma filial (Apurações PAI + Feedbacks).
 */
function buscarHistoricoCompletoClima(filial) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const userObj = obterUsuarioLogado(emailLogado);
    const filialNorm = normalizarFilialId(filial);
    const contatos = obterContatosPorFilial(filialNorm);

    // BLINDAGEM DE ACESSO
    if (contatos) {
      if (!usuarioPodeVerFilial(userObj, contatos.regional, contatos.diretoria)) {
        return [{ erro: 'ACESSO RESTRITO: A Filial ' + filialNorm + ' pertence à regional "' + contatos.regional + '", que está fora da sua alçada de visão.' }];
      }
    }

    var eventos = [];

    var dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      for (var i = 1; i < dataApuracao.length; i++) {
        var fVal = dataApuracao[i][3] ? dataApuracao[i][3].toString().trim() : '';
        if (normalizarFilialId(fVal) === filialNorm) {
          var dtStr = dataApuracao[i][1];
          if (dtStr instanceof Date) {
            dtStr = dtStr.toLocaleDateString('pt-BR');
          }

          var den1 = dataApuracao[i][23] ? dataApuracao[i][23].toString() : '';
          var den2 = dataApuracao[i][26] ? dataApuracao[i][26].toString() : '';
          var den3 = dataApuracao[i][29] ? dataApuracao[i][29].toString() : '';
          var denunciados = [den1, den2, den3].filter(function(n) { return n; }).join(', ');

          eventos.push({
            tipo: 'PAI',
            id: dataApuracao[i][0] ? dataApuracao[i][0].toString() : '',
            data: dtStr ? dtStr.toString() : 'N/A',
            titulo: 'Relatório de Apuração (PAI)',
            conclusao: dataApuracao[i][7] ? dataApuracao[i][7].toString() : 'Concluído',
            feedbackGerente: dataApuracao[i][17] ? dataApuracao[i][17].toString() : '',
            envolvidos: denunciados || 'Não informados',
            linkDoc: dataApuracao[i][8] ? dataApuracao[i][8].toString() : ''
          });
        }
      }
    }

    var dataInt = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataInt && dataInt.length > 1) {
      for (var j = 1; j < dataInt.length; j++) {
        var fValInt = dataInt[j][1] ? dataInt[j][1].toString().trim() : '';
        if (normalizarFilialId(fValInt) === filialNorm) {
          var dtStrInt = dataInt[j][0];
          if (dtStrInt instanceof Date) {
            dtStrInt = dtStrInt.toLocaleDateString('pt-BR');
          }

          var col1 = dataInt[j][8] ? dataInt[j][8].toString() : '';
          var col2 = dataInt[j][11] ? dataInt[j][11].toString() : '';
          var col3 = dataInt[j][14] ? dataInt[j][14].toString() : '';
          var colabs = [col1, col2, col3].filter(function(n) { return n; }).join(', ');

          eventos.push({
            tipo: 'FEEDBACK',
            id: dataInt[j][0] ? dataInt[j][0].toString() : '',
            data: dtStrInt ? dtStrInt.toString() : 'N/A',
            titulo: 'Acompanhamento / Feedback de Clima',
            status: dataInt[j][3] ? dataInt[j][3].toString() : 'Registrado',
            resumo: dataInt[j][4] ? dataInt[j][4].toString() : 'N/A',
            envolvidos: colabs || 'Equipe Geral / Loja',
            notaDepois: dataInt[j][5] ? dataInt[j][5].toString() : '',
            linkDoc: dataInt[j][2] ? dataInt[j][2].toString() : ''
          });
        }
      }
    }

    return eventos;
  } catch (e) {
    Logger.log("Erro ao buscar historico de clima: " + e.message);
    return [];
  }
}

/**
 * Consulta casos recentes registrados por filial para vinculo no formulário.
 */
function buscarCasosRecentesPorFilial(filial) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const userObj = obterUsuarioLogado(emailLogado);
    var filialNorm = normalizarFilialId(filial);
    const contatos = obterContatosPorFilial(filialNorm);

    // BLINDAGEM DE ACESSO
    if (contatos) {
      if (!usuarioPodeVerFilial(userObj, contatos.regional, contatos.diretoria)) {
        return [{ erro: 'ACESSO RESTRITO' }];
      }
    }

    var resultados = [];
    var dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      var headers = dataApuracao[0].map(function(h) { return normalizarTexto(h); });
      var colFilial = headers.indexOf('filial') !== -1 ? headers.indexOf('filial') : 3;
      var colDate = headers.indexOf('data registro') !== -1 ? headers.indexOf('data registro') : 1;
      var colConclusao = headers.indexOf('conclusao') !== -1 ? headers.indexOf('conclusao') : 7;
      var colApurador = headers.indexOf('apurador') !== -1 ? headers.indexOf('apurador') : 6;
      var colLink = headers.indexOf('link doc') !== -1 ? headers.indexOf('link doc') : 8;

      for (var i = 1; i < dataApuracao.length; i++) {
        var fVal = dataApuracao[i][colFilial] ? dataApuracao[i][colFilial].toString().trim() : '';
        if (normalizarFilialId(fVal) === filialNorm) {
          var dataStr = '';
          if (dataApuracao[i][colDate] instanceof Date) {
            dataStr = dataApuracao[i][colDate].toLocaleDateString('pt-BR');
          } else {
            dataStr = dataApuracao[i][colDate] ? dataApuracao[i][colDate].toString() : 'N/A';
          }

          var d1 = dataApuracao[i][23] ? dataApuracao[i][23].toString() : '';
          var d2 = dataApuracao[i][26] ? dataApuracao[i][26].toString() : '';
          var d3 = dataApuracao[i][29] ? dataApuracao[i][29].toString() : '';
          var consolidadoDenunciados = [d1, d2, d3].filter(function(n) { return n; }).join(', ') || 'N/A';
          var link = dataApuracao[i][colLink] ? dataApuracao[i][colLink].toString() : '';

          resultados.push({
            dataRegistro: dataStr,
            conclusao: dataApuracao[i][colConclusao] ? dataApuracao[i][colConclusao].toString() : 'Pendente',
            apurador: dataApuracao[i][colApurador] ? dataApuracao[i][colApurador].toString() : 'GP',
            denunciados: consolidadoDenunciados,
            linkDoc: link,
            regional: contatos ? contatos.regional : ''
          });
        }
      }
    }

    return resultados;
  } catch (err) {
    Logger.log("Erro ao buscar casos recentes por filial: " + err.message);
    return [];
  }
}

/**
 * Grava ou atualiza uma intervenção de feedback no banco de dados.
 */
function processarNovaIntervencao(dados) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    var userObj = obterUsuarioLogado(emailLogado);
    
    // TRAVA ABSOLUTA DE GRAVAÇÃO: Apenas GTGP, Coordenadores, Diretor RH e Master
    var cargoNorm = userObj.cargo ? userObj.cargo.toLowerCase() : '';
    var ehPermitido = false;
    
    if (userObj.role === 'MASTER') {
        ehPermitido = true;
    } else if (cargoNorm.includes('coord') || cargoNorm.includes('diretor rh') || cargoNorm.includes('diretorrh') || cargoNorm.includes('gtgp') || cargoNorm.includes('gp')) {
        if (userObj.role !== 'LIDERANCA' && userObj.role !== 'COMPLIANCE' && userObj.role !== 'GERENTE_LOJA') {
            ehPermitido = true;
        }
    }

    if (!ehPermitido) {
        return { sucesso: false, erro: 'Acesso Negado: Apenas a gestão do GP (Coordenadores, Diretor RH e GTGP) possui permissão para lançar ou editar feedbacks de clima.' };
    }

    var filial = normalizarFilialId(dados.filial);
    var contatos = obterContatosPorFilial(filial);

    // BLINDAGEM DE GRAVAÇÃO CROSS-REGIONAL
    if (contatos && !usuarioPodeVerFilial(userObj, contatos.regional, contatos.diretoria)) {
        return { sucesso: false, erro: 'Acesso Negado: Você não tem permissão para registrar ou editar clima em uma regional fora da sua alçada (' + contatos.regional + ').' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Intervencoes_Feedback') || ss.insertSheet('Intervencoes_Feedback');
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'ID Intervencao', 'Filial', 'Relatório PAI (Link)', 'Status Evolução', 'Detalhes Atividades', 'Nota Humor Depois', 'Nova Data Programada',
        'ID Colaborador 1', 'Nome Colaborador 1', 'Filial Colaborador 1',
        'ID Colaborador 2', 'Nome Colaborador 2', 'Filial Colaborador 2',
        'ID Colaborador 3', 'Nome Colaborador 3', 'Filial Colaborador 3',
        'Email Criador', 'ID Gerente Alvo', 'Nome Gerente Alvo'
      ]);
    }

    var idIntervencao = dados.id || '';
    var rowIndex = -1;

    if (idIntervencao) {
      var vals = sheet.getDataRange().getValues();
      for (var i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idIntervencao.trim().toUpperCase()) {
          var criadorOriginal = vals[i][16] ? vals[i][16].toString().toLowerCase().trim() : '';
          if (criadorOriginal && criadorOriginal !== emailLogado && userObj.role !== 'MASTER') {
             // Coordenadores podem editar os registros da equipe deles? 
             // Se sim, removemos essa trava. Mas manter a trava do criador evita sobrescritas acidentais.
             // Como a instrução é permitir gestão, liberamos a edição para coordenadores da mesma regional.
             if (!cargoNorm.includes('coord') && !cargoNorm.includes('gtgp')) {
                 return { sucesso: false, erro: 'Acesso Negado: Apenas a coordenação ou o criador original pode editar este registro.' };
             }
          }
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (!idIntervencao || rowIndex === -1) {
      idIntervencao = 'INT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    var colab1 = dados.colaboradores && dados.colaboradores[0] ? dados.colaboradores[0] : {};
    var colab2 = dados.colaboradores && dados.colaboradores[1] ? dados.colaboradores[1] : {};
    var colab3 = dados.colaboradores && dados.colaboradores[2] ? dados.colaboradores[2] : {};

    var rowData = [
      idIntervencao, filial, dados.linkDoc || 'Feedback Avulso / Orientativo (Sem PAI)', 
      dados.status_evolucao, dados.detalhes_intervencao, dados.nota_humor_depois || '', dados.data_intervencao || 'N/A',
      colab1.id || '', colab1.nome || '', colab1.filial || '',
      colab2.id || '', colab2.nome || '', colab2.filial || '',
      colab3.id || '', colab3.nome || '', colab3.filial || '',
      emailLogado, dados.gerente_alvo_id || '', dados.gerente_alvo_nome || ''
    ];

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { sucesso: true, id: idIntervencao };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}

/**
 * Consulta dados de feedback pendente para exibição na tela do Gerente de Loja.
 */
function buscarDadosFeedbackGerente(idFeedback) {
  try {
    if (!idFeedback) return { erro: 'ID do feedback não fornecido.' };
    var data = getCachedSheetData(SPREADSHEET_ID, 'Feedbacks_Gerentes');
    if (!data) return { erro: 'Base de feedbacks não localizada.' };
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === idFeedback.toString().trim().toUpperCase()) {
        return { 
          id: idFeedback, 
          filial: normalizarFilialId(data[i][1]), 
          feedback: data[i][3] ? data[i][3].toString() : '',
          status: data[i][5] ? data[i][5].toString() : 'Pendente'
        };
      }
    }
    return { erro: 'Plano de ação / Feedback ' + idFeedback + ' não localizado.' };
  } catch (err) {
    return { erro: 'Erro ao carregar feedback: ' + err.toString() };
  }
}

/**
 * Processa a resposta do gerente, emite Termo em PDF no Drive e marca status Respondido.
 */
function processarRespostaGerente(idFeedback, consideracoes, arquivosBase64) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Feedbacks_Gerentes');
    if (!sheet) return { sucesso: false, erro: 'Aba Feedbacks_Gerentes não encontrada.' };

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var filial = '';
    var feedbackOriginal = '';
    var nomeGerOriginal = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === idFeedback.toString().trim().toUpperCase()) { 
        rowIndex = i + 1; 
        filial = normalizarFilialId(data[i][1]); 
        feedbackOriginal = data[i][3] ? data[i][3].toString() : '';
        nomeGerOriginal = data[i][9] ? data[i][9].toString() : '';
        break; 
      }
    }
    if (rowIndex === -1) return { sucesso: false, erro: 'Feedback não localizado.' };

    var emailLogado = Session.getActiveUser().getEmail();
    var nomeLogado = obterUsuarioLogado(emailLogado).nome;
    var nomeFinalGestor = nomeGerOriginal || nomeLogado;
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var linksAnexos = [];
    var timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyyy_HHmm");

    var docTemp = DocumentApp.create("Termo_Ciente_F" + filial + "_" + timeStamp);
    var body = docTemp.getBody();
    
    body.appendParagraph("PLANO DE AÇÃO E TERMO DE CIÊNCIA DE FEEDBACK").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Filial: " + filial + " | Gestor: " + nomeFinalGestor);
    body.appendParagraph("Data da Confirmação: " + new Date().toLocaleString('pt-BR'));
    body.appendParagraph("\n1. DIRETRIZES DO GP:\n" + feedbackOriginal);
    body.appendParagraph("\n2. AÇÕES ADOTADAS PELO GESTOR:\n" + consideracoes);
    docTemp.saveAndClose();
    
    var pdfBlob = docTemp.getAs('application/pdf');
    var pdfFile = folder.createFile(pdfBlob).setName("Termo_Ciente_F" + filial + "_" + timeStamp + ".pdf");
    
    pdfFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    DriveApp.getFileById(docTemp.getId()).setTrashed(true);
    
    linksAnexos.push(pdfFile.getUrl());

    if (arquivosBase64 && arquivosBase64.length > 0) {
      arquivosBase64.forEach(function(arq) {
        if (arq.dados && arq.dados.indexOf(',') !== -1) {
          var split = arq.dados.split(',');
          var match = split[0].match(/:(.*?);/);
          var mime = match ? match[1] : 'image/png';
          var blob = Utilities.newBlob(Utilities.base64Decode(split[1]), mime, "Comprovante_F" + filial + "_" + (arq.nome || 'Anexo.png'));
          var fEv = folder.createFile(blob);
          fEv.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
          linksAnexos.push(fEv.getUrl());
        }
      });
    }

    sheet.getRange(rowIndex, 6).setValue('Respondido');
    sheet.getRange(rowIndex, 7).setValue(new Date());
    sheet.getRange(rowIndex, 8).setValue(consideracoes);
    sheet.getRange(rowIndex, 9).setValue(linksAnexos.join(' \n'));

    return { sucesso: true };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dispara e-mail contendo as diretrizes de feedback de clima para o gerente da filial.
 */
function enviarEmailGerenteFeedback(emailGerente, idFeedback, feedbackText, linkAcesso, filial, emailCopia) {
  var assunto = "[AÇÃO DE CLIMA E FEEDBACK] Diretrizes e Plano de Ação - Filial " + filial;
  var htmlBody = 
    '<div style="font-family: Arial, sans-serif; color: #1a1a1a; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">' +
      '<div style="background-color: #FF5A00; padding: 24px; color: white; text-align: center;">' +
        '<div style="background-color: white; color: #FF5A00; height: 40px; width: 40px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; margin-bottom: 10px;">M</div>' +
        '<h2 style="margin: 0; font-size: 20px; font-weight: 800; text-transform: uppercase;">Magazine Luiza</h2>' +
        '<p style="margin: 4px 0 0 0; font-size: 12px; font-weight: bold; opacity: 0.9;">PLANO DE DESENVOLVIMENTO E AÇÃO (FEEDBACK)</p>' +
      '</div><div style="padding: 24px; background-color: white;">' +
        '<p style="font-size: 15px;">Prezado(a) <strong>Gerente da Filial ' + filial + '</strong>,</p>' +
        '<p style="font-size: 14px; color: #4A5568;">Como parte das ações estruturadas da área de Gestão de Pessoas (GP), compartilhamos as seguintes diretrizes de clima para ciente e execução na sua unidade:</p>' +
        '<div style="background-color: #FFF5F5; border-left: 4px solid #FF5A00; padding: 18px; border-radius: 8px; margin: 20px 0; color: #C53030; font-style: italic;">"' + feedbackText + '"</div>' +
        '<p style="font-size: 14px; color: #4A5568;">Para formalizar o recebimento destas diretrizes e registrar as ações adotadas em loja, acesse o painel corporativo clicando abaixo.</p>' +
        '<div style="text-align: center; margin: 30px 0;"><a href="' + linkAcesso + '" target="_blank" rel="noopener noreferrer" style="background-color: #FF5A00; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Visualizar Diretrizes e Dar Ciente</a></div>' +
      '</div></div>';
  
  var emailsTo = emailGerente.split(';').join(',');
  var opcoes = { to: emailsTo, subject: assunto, htmlBody: htmlBody };
  if (emailCopia) {
    opcoes.cc = emailCopia;
  }
  MailApp.sendEmail(opcoes);
}
