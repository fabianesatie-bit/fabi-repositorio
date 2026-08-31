/**
 * SERVIÇO DE PROCESSAMENTO E CONSOLIDAÇÃO DE DADOS DAS PLANILHAS COM CACHE EM MEMÓRIA (CHUNKS)
 */

function putLargeCache(cache, key, valueString, ttlSeconds) {
  try {
    var chunkSize = 85000;
    var totalChunks = Math.ceil(valueString.length / chunkSize);
    cache.put(key + '_META', JSON.stringify({ chunks: totalChunks }), ttlSeconds);
    for (var i = 0; i < totalChunks; i++) {
      var chunk = valueString.substring(i * chunkSize, (i + 1) * chunkSize);
      cache.put(key + '_C_' + i, chunk, ttlSeconds);
    }
  } catch (e) {
    Logger.log('Erro ao salvar em chunk cache: ' + e.toString());
  }
}

function getLargeCache(cache, key) {
  try {
    var metaStr = cache.get(key + '_META');
    if (!metaStr) return null;
    var meta = JSON.parse(metaStr);
    var result = '';
    for (var i = 0; i < meta.chunks; i++) {
      var chunk = cache.get(key + '_C_' + i);
      if (chunk === null) return null;
      result += chunk;
    }
    return result;
  } catch (e) {
    Logger.log('Erro ao ler chunk cache: ' + e.toString());
    return null;
  }
}

function removeLargeCache(cache, key) {
  try {
    var metaStr = cache.get(key + '_META');
    if (metaStr) {
      var meta = JSON.parse(metaStr);
      for (var i = 0; i < meta.chunks; i++) {
        cache.remove(key + '_C_' + i);
      }
      cache.remove(key + '_META');
    }
    cache.remove(key);
  } catch (e) {}
}

function obterDadosAuditoria(forceRefresh) {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'APONTAMENTO_DATA_CACHE_V5';

    // 1. Tenta carregar dados consolidados do Cache em Chunks se não for refresh forçado
    if (!forceRefresh) {
      var cachedData = getLargeCache(cache, cacheKey);
      if (cachedData) {
        return cachedData;
      }
    }

    var mapaLojas = {};
    var listaLojas = [];
    var mapaHierarquia = {};

    // 2. Carregar DADOS_LOJAS da planilha DB_MASTER (com limites estritos de linha/coluna)
    try {
      var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
      var sheetLojas = ssMaster.getSheetByName(TAB_NAMES.LOJAS);
      if (sheetLojas) {
        var lastRowL = sheetLojas.getLastRow();
        var lastColL = sheetLojas.getLastColumn();
        if (lastRowL > 1 && lastColL > 0) {
          var dataLojas = sheetLojas.getRange(1, 1, lastRowL, lastColL).getValues();
          var h = dataLojas[0].map(function(x) { return String(x).toLowerCase().trim(); });
          var idxFid = h.findIndex(function(x) { return x.includes('filial') || x.includes('id'); });
          var idxNome = h.findIndex(function(x) { return x.includes('fantasia') || x.includes('nome'); });
          var idxReg = h.findIndex(function(x) { return x.includes('regional'); });
          var idxDir = h.findIndex(function(x) { return x.includes('diretoria') || x.includes('diretor'); });
          var idxEmail = h.findIndex(function(x) { return x.includes('gerente') || x.includes('email'); });
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
            var cid = idxCid !== -1 && dataLojas[i][idxCid] ? String(dataLojas[i][idxCid]).trim() : 'Franca';
            var uf = idxUf !== -1 && dataLojas[i][idxUf] ? String(dataLojas[i][idxUf]).trim() : 'SP';

            var objLoja = {
              filial: fNorm,
              nomeLoja: nome,
              regional: reg,
              diretoria: dir,
              gerenteEmail: email,
              cidade: cid,
              estado: uf
            };

            mapaLojas[fNorm] = objLoja;
            listaLojas.push(objLoja);

            if (!mapaHierarquia[dir]) mapaHierarquia[dir] = {};
            if (!mapaHierarquia[dir][reg]) mapaHierarquia[dir][reg] = [];
            mapaHierarquia[dir][reg].push(objLoja);
          }
        }
      }
    } catch (errLojas) {
      Logger.log('Erro ao ler DADOS_LOJAS: ' + errLojas.toString());
    }

    // 3. Carregar Apontamentos das Abas na Planilha de Auditoria
    var apontamentosBrutos = [];
    var totalIrregularidades = 0;
    var colabsUnicos = {};

    try {
      var ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
      var sheets = ssAuditoria.getSheets();

      sheets.forEach(function(sheet) {
        var sName = sheet.getName().toLowerCase();
        if (sName.includes('comprovantes')) return;

        var lastRowA = sheet.getLastRow();
        var lastColA = sheet.getLastColumn();
        if (lastRowA <= 1 || lastColA <= 0) return;

        var data = sheet.getRange(1, 1, lastRowA, lastColA).getValues();

        var h = data[0].map(function(x) { return String(x).toLowerCase().trim(); });
        var idxFid = h.findIndex(function(x) { return x === 'unidai' || x === 'unidade' || x.includes('filial') || x.includes('unidai') || x.includes('unidade') || x.includes('loja'); });
        var idxNome = h.findIndex(function(x) { return x === 'nome' || x === 'colaborador' || x.includes('colaborador') || (x.includes('nome') && !x.includes('regional')); });
        var idxCargo = h.findIndex(function(x) { return x.includes('cargo') || x.includes('funcao') || x.includes('função'); });
        var idxChapa = h.findIndex(function(x) { return x === 'cdi' || x.includes('chapa') || x.includes('cdi') || x.includes('contratado') || x.includes('matricula'); });
        var idxIrreg = h.findIndex(function(x) { return x.includes('irregularidade') || x.includes('tipo') || x.includes('documento'); });

        var isForaJornada = sName.includes('acesso') || sName.includes('jornada');
        var isHoraExtra = sName.includes('hora') || sName.includes('extra');
        var isBritanico = sName.includes('brit') || sName.includes('ajuste');

        if (!isForaJornada && !isHoraExtra && !isBritanico) return;

        for (var i = 1; i < data.length; i++) {
          var rawF = data[i][idxFid !== -1 ? idxFid : 0];
          var fNorm = normalizarFilialId(rawF);
          if (!fNorm) continue;

          var infoLoja = mapaLojas[fNorm] || {
            nomeLoja: 'Filial ' + fNorm,
            regional: 'Regional SP Interior',
            diretoria: 'Diretoria Lojas Sul/Sudeste'
          };

          var colabNome = idxNome !== -1 && data[i][idxNome] ? String(data[i][idxNome]).trim() : ('Colaborador ' + i);
          if (!colabNome || colabNome === 'Colaborador') continue;

          var tipo = idxIrreg !== -1 && data[i][idxIrreg] ? String(data[i][idxIrreg]).trim() : (isForaJornada ? 'Acesso Fora da Jornada' : (isHoraExtra ? 'Horas Extras Não Autorizadas' : 'Ajuste Britânico'));
          var chapa = idxChapa !== -1 && data[i][idxChapa] ? String(data[i][idxChapa]).trim() : ('78' + (1000 + i));
          var cargo = idxCargo !== -1 && data[i][idxCargo] ? String(data[i][idxCargo]).trim() : 'Vendedor';

          colabsUnicos[chapa || colabNome] = true;
          totalIrregularidades++;

          apontamentosBrutos.push({
            id: 'APONT-' + i,
            filialId: fNorm,
            filialNome: infoLoja.nomeLoja,
            regional: infoLoja.regional,
            diretoria: infoLoja.diretoria,
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

    // 4. Carregar Certificados Salvos
    var certificados = obterCertificadosTreinamento();

    var payloadResult = JSON.stringify({
      sucesso: true,
      kpisGlobais: {
        colaboradoresIrregulares: Object.keys(colabsUnicos).length,
        apontamentosCriticos: totalIrregularidades
      },
      filiaisAlertas: listaLojas,
      apontamentosBrutos: apontamentosBrutos,
      certificados: certificados,
      mapaHierarquia: mapaHierarquia
    });

    // 5. Salva no cache em chunks (suporta tamanhos > 100KB)
    putLargeCache(cache, cacheKey, payloadResult, CACHE_TTL_SECONDS);

    return payloadResult;
  } catch (err) {
    Logger.log('Erro em obterDadosAuditoria: ' + err.toString());
    return JSON.stringify({ sucesso: false, erro: err.toString() });
  }
}

function saveCertificadoTreinamento(payload) {
  try {
    if (!payload) throw new Error('Dados do certificado inválidos.');

    var ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    var sheet = ss.getSheetByName(TAB_NAMES.CERTIFICADOS);

    if (!sheet) {
      sheet = ss.insertSheet(TAB_NAMES.CERTIFICADOS);
      var headers = ['ID', 'Data_Registro', 'Filial_ID', 'Filial_Nome', 'Chapa', 'Colaborador_Nome', 'Cargo', 'Tipo_Treinamento', 'Link_Certificado', 'Status', 'Registrado_Por', 'Observacoes'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#0086ff').setFontColor('#ffffff').setFontWeight('bold');
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
      String(payload.tipoTreinamento || 'Ponto Eletrônico - O Guia Essencial do Colaborador'),
      String(payload.linkComprovante || ''),
      String(payload.status || 'Aprovado'),
      String(payload.registradoPor || 'Gerente de Loja'),
      String(payload.observacoes || '')
    ]);

    // Invalida o cache após salvar novo certificado
    removeLargeCache(CacheService.getScriptCache(), 'AUDITORIA_DATA_CACHE_V3');

    return { success: true, message: 'Certificado registrado com sucesso na aba Comprovantes_Treinamento!' };
  } catch (err) {
    Logger.log('Erro ao salvar certificado: ' + err.toString());
    return { success: false, error: err.message };
  }
}

function obterCertificadosTreinamento() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    var sheet = ss.getSheetByName(TAB_NAMES.CERTIFICADOS);
    if (!sheet) return [];

    var lastRowC = sheet.getLastRow();
    var lastColC = sheet.getLastColumn();
    if (lastRowC <= 1 || lastColC <= 0) return [];

    var data = sheet.getRange(1, 1, lastRowC, lastColC).getValues();

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

/**
 * UPLOAD DE CERTIFICADO DO COMPUTADOR PARA O GOOGLE DRIVE E INSERÇÃO NA PLANILHA
 */
function uploadCertificadoFile(payload) {
  try {
    if (!payload || !payload.fileData) {
      throw new Error('Nenhum arquivo enviado.');
    }

    var filialId = String(payload.filialId || '');
    var filialNome = String(payload.filialNome || ('Filial ' + filialId));
    var chapa = String(payload.chapa || '');
    var colabNome = String(payload.colaboradorNome || 'Colaborador');
    var cargo = String(payload.cargo || 'Vendedor');
    var curso = String(payload.tipoTreinamento || 'Ponto Eletrônico - O Guia Essencial do Colaborador');
    var obs = String(payload.observacoes || '');

    // 1. Obter ou criar pasta no Google Drive para os certificados
    var folderName = 'Certificados_Treinamento_GP360';
    var targetFolder;
    try {
      var folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        targetFolder = folders.next();
      } else {
        targetFolder = DriveApp.createFolder(folderName);
      }
    } catch (eDriveFolder) {
      targetFolder = DriveApp.getRootFolder();
    }

    // 2. Converter Base64 do cliente para Blob
    var base64Data = payload.fileData.split(',')[1] || payload.fileData;
    var ext = payload.extension || 'pdf';
    var fileName = 'Certificado_' + filialId + '_' + chapa + '_' + new Date().getTime() + '.' + ext;
    var contentType = payload.mimeType || 'application/pdf';
    var decodedBytes = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decodedBytes, contentType, fileName);

    // 3. Salvar no Drive e ajustar permissão de leitura
    var driveFile = targetFolder.createFile(blob);
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eSharing) {
      Logger.log('Aviso ao ajustar compartilhamento do Drive: ' + eSharing.toString());
    }
    var fileUrl = driveFile.getUrl();

    // 4. Salvar registro na planilha
    var certPayload = {
      filialId: filialId,
      filialNome: filialNome,
      chapa: chapa,
      colaboradorNome: colabNome,
      cargo: cargo,
      tipoTreinamento: curso,
      linkComprovante: fileUrl,
      status: 'Aprovado',
      registradoPor: payload.registradoPor || 'Gerente de Loja (Portal Exclusivo)',
      observacoes: obs
    };

    var resultSave = saveCertificadoTreinamento(certPayload);
    if (!resultSave.success) throw new Error(resultSave.error);

    return {
      success: true,
      fileUrl: fileUrl,
      message: 'Certificado do colaborador ' + colabNome + ' enviado e registrado com sucesso!'
    };
  } catch (err) {
    Logger.log('Erro em uploadCertificadoFile: ' + err.toString());
    return { success: false, error: err.message || err.toString() };
  }
}

/**
 * RETORNA A LISTA DE COLABORADORES PENDENTES DE UMA FILIAL ESPECÍFICA (DEDUPLICADA POR CHAPA)
 */
function obterColaboradoresPendentesFilial(filialId) {
  try {
    var rawData = obterDadosAuditoria(false);
    var data = typeof rawData === 'string' ? JSON.parse(rawData) : rawData;
    if (!data.sucesso) return { sucesso: false, erro: data.erro };

    var fNorm = normalizarFilialId(filialId);
    var pendentes = (data.apontamentosBrutos || []).filter(function(a) {
      return normalizarFilialId(a.filialId) === fNorm || String(a.filialId) === String(filialId);
    });

    var certs = data.certificados || [];
    var certsMap = {};
    certs.forEach(function(c) {
      if (normalizarFilialId(c.filialId) === fNorm || String(c.filialId) === String(filialId)) {
        certsMap[c.chapa] = c;
      }
    });

    // Agrupar colaboradores por chapa para exibir 1 única linha por funcionário
    var mapaColabs = {};
    pendentes.forEach(function(p) {
      var key = p.chapa || p.nome;
      if (!mapaColabs[key]) {
        mapaColabs[key] = {
          filialId: p.filialId,
          filialNome: p.filialNome,
          chapa: p.chapa,
          nome: p.nome,
          cargo: p.cargo,
          irregularidadesMap: {},
          totalOcorrencias: 0
        };
      }
      var tipo = p.tipoIrregularidade || 'Apontamento em Aberto';
      mapaColabs[key].irregularidadesMap[tipo] = (mapaColabs[key].irregularidadesMap[tipo] || 0) + 1;
      mapaColabs[key].totalOcorrencias += (p.quantidadeMes || 1);
    });

    var resultado = Object.keys(mapaColabs).map(function(k) {
      var c = mapaColabs[k];
      var certEnviado = certsMap[c.chapa];
      var tiposFormatados = Object.keys(c.irregularidadesMap).map(function(t) {
        var qtd = c.irregularidadesMap[t];
        return qtd > 1 ? (t + ' (' + qtd + 'x)') : t;
      }).join(', ');

      return {
        filialId: c.filialId,
        filialNome: c.filialNome,
        chapa: c.chapa,
        nome: c.nome,
        cargo: c.cargo,
        tipoIrregularidade: tiposFormatados,
        totalOcorrencias: c.totalOcorrencias,
        concluido: !!certEnviado,
        linkCertificado: certEnviado ? certEnviado.linkComprovante : ''
      };
    });

    var infoLoja = (data.filiaisAlertas || []).find(function(l) { 
      return normalizarFilialId(l.filial) === fNorm || String(l.filial) === String(filialId); 
    }) || {
      filial: fNorm,
      nomeLoja: 'Filial ' + fNorm
    };

    return {
      sucesso: true,
      loja: infoLoja,
      colaboradores: resultado
    };
  } catch (err) {
    Logger.log('Erro em obterColaboradoresPendentesFilial: ' + err.toString());
    return { sucesso: false, erro: err.toString() };
  }
}
