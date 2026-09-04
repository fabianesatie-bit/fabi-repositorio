/**
 * SERVIÇO DE PROCESSAMENTO, CONSOLIDAÇÃO DE DADOS E DISPARO DE NOTIFICAÇÕES
 * ECOSSISTEMA GP360 - MÓDULO APONTAMENTO
 */

function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

var Service_Data = (function() {

  function obterDadosAuditoria(forceRefresh) {
    try {
      var cache = CacheService.getScriptCache();
      var cacheKey = 'AUDITORIA_DATA_CACHE_V21';
      
      if (!forceRefresh) {
        var cachedData = cache.get(cacheKey);
        if (cachedData) {
          return cachedData;
        }
      }

      var mapaLojas = {};
      var listaLojas = [];
      var mapaHierarquia = {};

      try {
        var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
        var sheetLojas = ssMaster.getSheetByName(TAB_NAMES.LOJAS);
        
        if (!sheetLojas) {
          var sheets = ssMaster.getSheets();
          for (var k = 0; k < sheets.length; k++) {
            var sn = sheets[k].getName().toLowerCase();
            if (sn.includes('loja') || sn.includes('dados')) {
              sheetLojas = sheets[k];
              break;
            }
          }
        }

        if (sheetLojas) {
          var dataLojas = sheetLojas.getDataRange().getValues();
          if (dataLojas.length > 1) {
            var h = dataLojas[0].map(function(x) { return String(x).toLowerCase().trim(); });
            var idxFid = h.findIndex(function(x) { return x.includes('filial') || x.includes('id'); });
            var idxNome = h.findIndex(function(x) { return x.includes('fantasia') || (x.includes('nome') && !x.includes('regional')); });
            var idxReg = h.findIndex(function(x) { return x.includes('regional'); });
            var idxDir = h.findIndex(function(x) { return x.includes('diretoria') || x.includes('diretor'); });
            var idxEmail = h.findIndex(function(x) { return x.includes('gerente') || x.includes('email'); });
            var idxTel = h.findIndex(function(x) { return x.includes('telefone') || x.includes('whatsapp') || x.includes('celular') || x.includes('contato') || x.includes('tel'); });

            var idxCid = h.findIndex(function(x) { return x.includes('cidade'); });
            var idxUf = h.findIndex(function(x) { return x.includes('estado') || x.includes('uf'); });

            for (var i = 1; i < dataLojas.length; i++) {
              var rawId = dataLojas[i][idxFid !== -1 ? idxFid : 0];
              var fNorm = normalizarFilialId(rawId);
              if (!fNorm) continue;

              var dir = idxDir !== -1 && dataLojas[i][idxDir] ? String(dataLojas[i][idxDir]).trim() : 'Diretoria Lojas Sul/Sudeste';
              var reg = idxReg !== -1 && dataLojas[i][idxReg] ? String(dataLojas[i][idxReg]).trim() : 'Regional SP Interior';
              var nome = idxNome !== -1 && dataLojas[i][idxNome] ? String(dataLojas[i][idxNome]).trim() : ('Filial ' + fNorm);
              var email = idxEmail !== -1 && dataLojas[i][idxEmail] ? String(dataLojas[i][idxEmail]).trim() : ('gerente.filial' + fNorm + '@magazineluiza.com.br');
              
              var valTelRaw = (dataLojas[i][7] !== undefined && dataLojas[i][7] !== null && String(dataLojas[i][7]).trim() !== '') 
                ? dataLojas[i][7] 
                : (idxTel !== -1 ? dataLojas[i][idxTel] : '');
                
              var tel = String(valTelRaw || '').replace(/\D/g, '').trim();
              
              var cid = idxCid !== -1 && dataLojas[i][idxCid] ? String(dataLojas[i][idxCid]).trim() : 'Franca';
              var uf = idxUf !== -1 && dataLojas[i][idxUf] ? String(dataLojas[i][idxUf]).trim() : 'SP';

              if (mapaLojas[fNorm]) {
                if (!tel && mapaLojas[fNorm].gerenteTelefone) {
                  tel = mapaLojas[fNorm].gerenteTelefone;
                }
              }

              var objLoja = {
                filial: fNorm,
                nomeLoja: nome,
                regional: reg,
                diretoria: dir,
                gerenteEmail: email,
                gerenteTelefone: tel,
                telefone: tel,
                tel: tel,
                cidade: cid,
                estado: uf
              };

              mapaLojas[fNorm] = objLoja;
              
              var idxExistente = listaLojas.findIndex(function(l) { return l.filial === fNorm; });
              if (idxExistente !== -1) {
                listaLojas[idxExistente] = objLoja;
              } else {
                listaLojas.push(objLoja);
              }

              if (!mapaHierarquia[dir]) mapaHierarquia[dir] = {};
              if (!mapaHierarquia[dir][reg]) mapaHierarquia[dir][reg] = [];
              if (!mapaHierarquia[dir][reg].some(function(l) { return l.filial === fNorm; })) {
                mapaHierarquia[dir][reg].push(objLoja);
              }
            }
          }
        }
      } catch (errLojas) {
        Logger.log('Erro ao ler DADOS_LOJAS: ' + errLojas.toString());
      }

      var apontamentosBrutos = [];
      var totalIrregularidades = 0;
      var colabsUnicos = {};

      try {
        var ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
        var sheetsAud = ssAuditoria.getSheets();

        sheetsAud.forEach(function(sheet) {
          var sName = sheet.getName().toLowerCase();
          if (sName.includes('comprovantes') || sName.includes('certificados')) return;

          var isForaJornada = sName.includes('acesso') || sName.includes('jornada');
          var isHoraExtra = sName.includes('hora') || sName.includes('extra');
          var isBritanico = sName.includes('brit') || sName.includes('ajuste');

          if (!isForaJornada && !isHoraExtra && !isBritanico) return;

          var data = sheet.getDataRange().getValues();
          if (data.length <= 1) return;

          var h = data[0].map(function(x) { return String(x).toLowerCase().trim(); });
          var idxFid = h.findIndex(function(x) { return x.includes('unidai') || x.includes('filial') || x.includes('unidade') || x.includes('loja'); });
          var idxNome = h.findIndex(function(x) { return x === 'nome' || x.includes('colaborador') || (x.includes('nome') && !x.includes('regional')); });
          if (idxNome === -1) idxNome = 7;

          var idxCargo = h.findIndex(function(x) { return x.includes('cargo') || x.includes('funcao') || x.includes('função'); });
          var idxChapa = h.findIndex(function(x) { return x === 'cdi' || x.includes('chapa') || x.includes('contratado') || x.includes('matricula'); });
          var idxIrreg = h.findIndex(function(x) { return x.includes('irregularidade') || x.includes('tipo') || x.includes('documento'); });

          for (var i = 1; i < data.length; i++) {
            var rawF = data[i][idxFid !== -1 ? idxFid : 0];
            var fNorm = normalizarFilialId(rawF);
            if (!fNorm) continue;

            var colabNome = idxNome !== -1 && data[i][idxNome] ? String(data[i][idxNome]).trim() : ('Colaborador ' + i);
            if (!colabNome || colabNome === 'Colaborador' || colabNome.toLowerCase().includes('regional')) continue;

            var tipo = idxIrreg !== -1 && data[i][idxIrreg] ? String(data[i][idxIrreg]).trim() : (isForaJornada ? 'Acesso Fora da Jornada' : (isHoraExtra ? 'Horas Extras Não Autorizadas' : 'Ajuste Britânico'));
            var chapa = idxChapa !== -1 && data[i][idxChapa] ? String(data[i][idxChapa]).trim() : ('78' + (1000 + i));
            var cargo = idxCargo !== -1 && data[i][idxCargo] ? String(data[i][idxCargo]).trim() : 'Vendedor';

            colabsUnicos[chapa || colabNome] = true;
            totalIrregularidades++;

            apontamentosBrutos.push({
              id: 'APONT-' + i,
              filialId: fNorm,
              chapa: chapa,
              nome: colabNome,
              cargo: cargo,
              tipoIrregularidade: tipo,
              quantidadeMes: 1
            });
          }
        });
      } catch (errAud) {
        Logger.log('Erro ao ler abas de auditoria: ' + errAud.toString());
      }

      var certificados = obterCertificadosTreinamento();

      var payloadObj = {
        sucesso: true,
        perfil: {
          nome: 'Coordenador GP',
          role: 'CoordenadorGP' // Atribui perfil Padrão
        },
        kpisGlobais: {
          colaboradoresIrregulares: Object.keys(colabsUnicos).length,
          apontamentosCriticos: totalIrregularidades
        },
        filiaisAlertas: listaLojas,
        mapaLojas: mapaLojas,
        apontamentosBrutos: apontamentosBrutos,
        certificados: certificados,
        mapaHierarquia: mapaHierarquia
      };

      var payloadResult = JSON.stringify(payloadObj);

      try {
        cache.put(cacheKey, payloadResult, 600);
      } catch (eCache) {
        Logger.log('Aviso ao salvar no cache: ' + eCache.toString());
      }

      return payloadResult;
    } catch (err) {
      Logger.log('Erro em obterDadosAuditoria: ' + err.toString());
      return JSON.stringify({ sucesso: false, erro: err.toString() });
    }
  }

  function salvarTelefoneGerente(filialId, novoTelefone) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(10000)) {
        return { success: false, error: 'Sistema ocupado. Tente novamente em instantes.' };
      }

      var fNorm = normalizarFilialId(filialId);
      if (!fNorm) throw new Error('ID de filial inválido.');

      var telLimp = String(novoTelefone || '').replace(/\D/g, '').trim();

      var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
      var sheetLojas = ssMaster.getSheetByName(TAB_NAMES.LOJAS);
      if (!sheetLojas) {
        var sheets = ssMaster.getSheets();
        for (var k = 0; k < sheets.length; k++) {
          var sn = sheets[k].getName().toLowerCase();
          if (sn.includes('loja') || sn.includes('dados')) {
            sheetLojas = sheets[k];
            break;
          }
        }
      }

      if (!sheetLojas) throw new Error('Aba DADOS_LOJAS não encontrada.');

      var data = sheetLojas.getDataRange().getValues();
      if (data.length <= 1) throw new Error('Aba DADOS_LOJAS vazia.');

      var h = data[0].map(function(x) { return String(x).toLowerCase().trim(); });
      var idxFid = h.findIndex(function(x) { return x.includes('filial') || x.includes('id'); });
      var idxTel = 7;

      var alterou = false;
      for (var i = 1; i < data.length; i++) {
        var rawF = data[i][idxFid !== -1 ? idxFid : 0];
        if (normalizarFilialId(rawF) === fNorm) {
          sheetLojas.getRange(i + 1, idxTel + 1).setValue(telLimp);
          alterou = true;
        }
      }

      if (alterou) {
        CacheService.getScriptCache().remove('AUDITORIA_DATA_CACHE_V21');
        return { success: true, message: 'Telefone salvo com sucesso em todas as linhas da filial!' };
      }

      return { success: false, error: 'Filial não encontrada na aba DADOS_LOJAS.' };
    } catch (err) {
      Logger.log('Erro ao salvar telefone do gerente: ' + err.toString());
      return { success: false, error: err.message };
    } finally {
      lock.releaseLock();
    }
  }

  function sendEmailRegional(regionalNome) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(15000)) {
        return { success: false, error: 'Outro envio está sendo processado. Aguarde alguns segundos.' };
      }

      var userEmail = Session.getActiveUser().getEmail() || 'coordenacao.gp@magazineluiza.com.br';
      var dadosAuditoriaRaw = obterDadosAuditoria(false);
      var dataObj = typeof dadosAuditoriaRaw === 'string' ? JSON.parse(dadosAuditoriaRaw) : dadosAuditoriaRaw;

      if (!dataObj || !dataObj.sucesso) throw new Error('Falha ao obter base de dados para envio.');

      var lojasFiltradas = (dataObj.filiaisAlertas || []).filter(function(l) {
        return regionalNome === 'TODAS' || l.regional === regionalNome;
      });

      var enviados = 0;
      var webUrl = ScriptApp.getService().getUrl();

      lojasFiltradas.forEach(function(loja) {
        if (!loja.gerenteEmail) return;

        var fid = normalizarFilialId(loja.filial);
        var aponts = (dataObj.apontamentosBrutos || []).filter(function(a) { return normalizarFilialId(a.filialId) === fid; });
        if (aponts.length === 0) return;

        var linkPortal = webUrl + '?mode=certificado&filialId=' + fid;
        var linkCurso = 'https://universidadeluiza.com.br/app/home/canal/magalu?section=csc-conte-sempre-comigo&trail=ponto-eletronico-o-guia-essencial-do-colaborador';

        var htmlBody = '<div style="font-family: Arial, sans-serif; color: #1E293B; max-width: 600px; margin: 0 auto; border: 1px solid #E2E8F0; border-radius: 14px; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.05);">' +
          '<div style="background: linear-gradient(135deg, #7C3AED, #2563EB); padding: 24px;">' +
            '<h2 style="color: #FFFFFF; margin: 0; font-size: 20px;">Acompanhamento de Ponto Eletrônico — GP360</h2>' +
            '<p style="color: #E0E7FF; margin: 5px 0 0 0; font-size: 13px;">' + loja.nomeLoja + ' (' + fid + ') • ' + loja.regional + '</p>' +
          '</div>' +
          '<div style="padding: 24px; background-color: #FFFFFF;">' +
            '<p style="font-size: 14px; margin-top: 0;">Olá, <b>Gerente de Loja</b>,</p>' +

            '<div style="background-color: #FFF1F2; border-left: 5px solid #F43F5E; padding: 14px 18px; margin: 16px 0 20px 0; border-radius: 8px;">' +
              '<p style="margin: 0; font-size: 12px; color: #9F1239; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;">⚠️ TERMO DE COMPLIANCE E CONFIDENCIALIDADE:</p>' +
              '<p style="margin: 6px 0 0 0; font-size: 11px; color: #881337; line-height: 1.5; font-weight: 500;">' +
                'Esta mensagem destina-se exclusivamente à gestão da loja. É <b>estritamente proibido</b> compartilhar este conteúdo em grupos de WhatsApp, redes sociais ou encaminhar listas de inconsistências diretamente ao colaborador. Faça o alinhamento presencial e forneça unicamente o link do curso.' +
              '</p>' +
            '</div>' +

            '<p style="font-size: 13px; color: #475569; line-height: 1.5;">Identificamos apontamentos críticos de jornada pendentes de regularização na sua unidade. Por favor, execute as ações operacionais abaixo:</p>' +

            '<div style="margin: 20px 0;">' +
              '<a href="' + linkCurso + '" target="_blank" style="display: block; background-color: #EFF6FF; border: 1px solid #BFDBFE; color: #1E40AF; padding: 12px; border-radius: 10px; text-decoration: none; font-size: 12px; font-weight: bold; text-align: center; margin-bottom: 10px;">1. Link do Treinamento para o Colaborador (Universidade Luiza)</a>' +
              '<a href="' + linkPortal + '" target="_blank" style="display: block; background-color: #7C3AED; color: #FFFFFF; padding: 12px; border-radius: 10px; text-decoration: none; font-size: 12px; font-weight: bold; text-align: center;">2. Portal do Gerente — Enviar Comprovantes</a>' +
            '</div>' +
            '<p style="font-size: 11px; color: #94A3B8; text-align: center; margin-bottom: 0;">Atenciosamente,<br><b>Coordenação de Gestão de Pessoas — GP360</b></p>' +
          '</div>' +
        '</div>';

        MailApp.sendEmail({
          to: loja.gerenteEmail,
          subject: ' [AÇÃO OBRIGATÓRIA] Regularização de Ponto — ' + loja.nomeLoja,
          htmlBody: htmlBody,
          replyTo: userEmail,
          name: 'Coordenação GP360'
        });

        enviados++;
      });

      return { success: true, message: 'Disparo concluído com sucesso! ' + enviados + ' e-mails enviados.' };
    } catch (err) {
      Logger.log('Erro em sendEmailRegional: ' + err.toString());
      return { success: false, error: err.message };
    } finally {
      lock.releaseLock();
    }
  }

  function obterColaboradoresPendentesFilial(filialId) {
    try {
      var fNorm = normalizarFilialId(filialId);
      if (!fNorm) return { sucesso: false, erro: 'ID da filial não especificado.' };

      var dadosAuditoriaRaw = obterDadosAuditoria(false);
      var dataObj = typeof dadosAuditoriaRaw === 'string' ? JSON.parse(dadosAuditoriaRaw) : dadosAuditoriaRaw;

      if (!dataObj || !dataObj.sucesso) return { sucesso: false, erro: 'Falha ao buscar base de dados.' };

      var infoLoja = dataObj.mapaLojas[fNorm] || { nomeLoja: 'Filial ' + fNorm, filial: fNorm };
      var certs = dataObj.certificados || [];

      var certsMap = {};
      certs.forEach(function(c) {
        if (normalizarFilialId(c.filialId) === fNorm) {
          certsMap[c.chapa] = c.linkComprovante || '#';
        }
      });

      var colabsMap = {};
      (dataObj.apontamentosBrutos || []).forEach(function(item) {
        if (normalizarFilialId(item.filialId) === fNorm) {
          var cKey = item.chapa || item.nome;
          if (!colabsMap[cKey]) {
            colabsMap[cKey] = {
              filialId: fNorm,
              filialNome: infoLoja.nomeLoja,
              chapa: item.chapa,
              nome: item.nome,
              cargo: item.cargo,
              tipoIrregularidade: item.tipoIrregularidade,
              concluido: !!certsMap[item.chapa],
              linkCertificado: certsMap[item.chapa] || ''
            };
          }
        }
      });

      var listaColabs = Object.keys(colabsMap).map(function(k) { return colabsMap[k]; });

      return {
        sucesso: true,
        loja: infoLoja,
        colaboradores: listaColabs
      };
    } catch (err) {
      return { sucesso: false, erro: err.toString() };
    }
  }

  function uploadCertificadoFile(payload) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(15000)) {
        return { success: false, error: 'O servidor está processando outro arquivo. Tente novamente em alguns segundos.' };
      }

      if (!payload || !payload.fileData) throw new Error('Dados do arquivo inválidos.');

      var fNorm = normalizarFilialId(payload.filialId);
      var folder;
      try {
        folder = DriveApp.getFolderById('1_FOLDER_CERTIFICADOS_GP360');
      } catch(e) {
        folder = DriveApp.getRootFolder();
      }

      var contentType = payload.mimeType || 'application/pdf';
      var base64Clean = payload.fileData.substr(payload.fileData.indexOf(',') + 1);
      var decodedData = Utilities.base64Decode(base64Clean);
      var fileName = 'CERT_' + fNorm + '_' + payload.chapa + '_' + new Date().getTime() + '.' + (payload.extension || 'pdf');
      
      var blob = Utilities.newBlob(decodedData, contentType, fileName);
      var fileCreated = folder.createFile(blob);
      fileCreated.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var fileUrl = fileCreated.getUrl();

      saveCertificadoTreinamento({
        filialId: fNorm,
        filialNome: payload.filialNome || ('Filial ' + fNorm),
        chapa: payload.chapa,
        colaboradorNome: payload.colaboradorNome,
        cargo: payload.cargo,
        tipoTreinamento: payload.tipoTreinamento || 'Ponto Eletrônico - Guia do Colaborador',
        linkComprovante: fileUrl,
        status: 'Aprovado',
        registradoPor: payload.registradoPor || 'Gerente de Loja (Portal)'
      });

      return { success: true, fileUrl: fileUrl, message: 'Certificado enviado e registrado com sucesso!' };
    } catch (err) {
      Logger.log('Erro ao fazer upload do certificado: ' + err.toString());
      return { success: false, error: err.message };
    } finally {
      lock.releaseLock();
    }
  }

  function saveCertificadoTreinamento(payload) {
    var lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(10000)) {
        return { success: false, error: 'O sistema está processando outra solicitação simultânea. Tente novamente em instantes.' };
      }

      if (!payload) throw new Error('Dados do certificado inválidos.');

      var ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
      var sheet = ss.getSheetByName(TAB_NAMES.CERTIFICADOS);

      if (!sheet) {
        sheet = ss.insertSheet(TAB_NAMES.CERTIFICADOS);
        var headers = ['ID', 'Data_Registro', 'Filial_ID', 'Filial_Nome', 'Chapa', 'Colaborador_Nome', 'Cargo', 'Tipo_Treinamento', 'Link_Certificado', 'Status', 'Registrado_Por', 'Observacoes'];
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setBackground('#7C3AED').setFontColor('#ffffff').setFontWeight('bold');
        sheet.setFrozenRows(1);
      }

      var dataAgora = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
      var idNovo = 'CERT-' + new Date().getTime();

      sheet.appendRow([
        idNovo,
        dataAgora,
        String(payload.filialId || ''),
        String(payload.filialNome || ''),
        String(payload.chapa || ''),
        String(payload.colaboradorNome || ''),
        String(payload.cargo || ''),
        String(payload.tipoTreinamento || 'Ponto Eletrônico - Guia do Colaborador'),
        String(payload.linkComprovante || ''),
        String(payload.status || 'Aprovado'),
        String(payload.registradoPor || 'Gestor GP'),
        String(payload.observacoes || '')
      ]);

      CacheService.getScriptCache().remove('AUDITORIA_DATA_CACHE_V21');

      return { success: true, message: 'Certificado de treinamento registrado com sucesso!' };
    } catch (err) {
      Logger.log('Erro ao salvar certificado: ' + err.toString());
      return { success: false, error: err.message };
    } finally {
      lock.releaseLock();
    }
  }

  function obterCertificadosTreinamento() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
      var sheet = ss.getSheetByName(TAB_NAMES.CERTIFICADOS);
      if (!sheet) return [];

      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [];

      var list = [];
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        list.push({
          id: String(data[i][0]),
          dataEnvio: String(data[i][1]),
          filialId: String(data[i][2]),
          filialNome: String(data[i][3]),
          chapa: String(data[i][4]),
          colaboradorNome: String(data[i][5]),
          cargo: String(data[i][6]),
          tipoTreinamento: String(data[i][7]),
          linkComprovante: String(data[i][8]),
          status: String(data[i][9] || 'Aprovado')
        });
      }
      return list;
    } catch (err) {
      return [];
    }
  }

  return {
    obterDadosAuditoria: obterDadosAuditoria,
    salvarTelefoneGerente: salvarTelefoneGerente,
    sendEmailRegional: sendEmailRegional,
    obterColaboradoresPendentesFilial: obterColaboradoresPendentesFilial,
    uploadCertificadoFile: uploadCertificadoFile,
    saveCertificadoTreinamento: saveCertificadoTreinamento,
    obterCertificadosTreinamento: obterCertificadosTreinamento
  };

})();
