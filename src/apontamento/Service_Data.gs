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
    var cacheKey = 'AUDITORIA_DATA_CACHE_V3';

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
        var idxFid = h.findIndex(function(x) { return x.includes('filial') || x.includes('unidade') || x.includes('loja'); });
        var idxNome = h.findIndex(function(x) { return x.includes('nome') || x.includes('colaborador'); });
        var idxCargo = h.findIndex(function(x) { return x.includes('cargo') || x.includes('funcao') || x.includes('função'); });
        var idxChapa = h.findIndex(function(x) { return x.includes('chapa') || x.includes('cdi') || x.includes('contratado') || x.includes('matricula'); });
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
