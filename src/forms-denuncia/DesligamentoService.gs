// =============================================================================
// FLUXO DE DESLIGAMENTOS E ALÇADAS DE APROVAÇÃO
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: DesligamentoService.gs
// =============================================================================

/**
 * Processa e cria um novo dossiê de desligamento com alçadas de aprovação.
 */
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
          const criadorOriginal = desData[i][21] ? desData[i][21].toString().toLowerCase().trim() : '';
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode editar este registro.' };
          }
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
          const match = splitData[0].match(/:(.*?);/);
          const contentType = match ? match[1] : 'image/png';
          const rawData = Utilities.base64Decode(splitData[1]);
          const fileName = "Resultados_F" + filial + "_" + dados.colaborador_id + "_" + index + "_" + (arq.nome || ".png");
          const imgFile = folder.createFile(Utilities.newBlob(rawData, contentType, fileName));
          imgFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
          linksImagens.push(imgFile.getUrl());
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
      aplicarPermissoesArquivo(newDocId, emailRegionalGP, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailDiretorOp, 'LEITOR');
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
    }

    return { sucesso: true, link: docLink, id: idDesligamento };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
