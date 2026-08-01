/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Carrega estado inicial completo com verificação de segurança, moedas, rankings e integrações
 */
function obterDadosIniciais(filtroMes, filtroAno) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) {
    return { temAcesso: false, motivo: controle.motivo };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Data de referência (hoje ou mês/ano filtrado)
  var agora = new Date();
  var mesAlvo = (filtroMes !== undefined && filtroMes !== null && filtroMes !== '') ? parseInt(filtroMes) : agora.getMonth() + 1;
  var anoAlvo = (filtroAno !== undefined && filtroAno !== null && filtroAno !== '') ? parseInt(filtroAno) : agora.getFullYear();

  // 1. Dicionário de Prêmios/Moedas
  var dicionarioPremios = {};
  var abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
  if (abaPremios) {
    var dadosP = abaPremios.getDataRange().getValues();
    for (var i = 1; i < dadosP.length; i++) {
      var natKey = String(dadosP[i][0]).trim().toUpperCase();
      var pts = parseFloat(dadosP[i][1]) || 0;
      dicionarioPremios[natKey] = pts;
    }
  }

  // Valores padrão para regras específicas
  if (!dicionarioPremios['CANAL']) dicionarioPremios['CANAL'] = 2; // Apurações Canal = 2
  if (!dicionarioPremios['INTERNA']) dicionarioPremios['INTERNA'] = 4; // Apurações Interna = 4
  if (!dicionarioPremios['ATENDIMENTO_SOCIAL']) dicionarioPremios['ATENDIMENTO_SOCIAL'] = 2;

  // 2. Leitura de Lojas e Regionais
  var lojas = [];
  var mapaLojaRegional = {};
  var abaLojas = ss.getSheetByName('DADOS_LOJAS');
  if (abaLojas) {
    var dadosLojas = abaLojas.getDataRange().getValues();
    for (var j = 1; j < dadosLojas.length; j++) {
      var fId = ("0000" + dadosLojas[j][0]).slice(-4);
      var reg = String(dadosLojas[j][2] || '').trim().toUpperCase();
      mapaLojaRegional[fId] = reg;
      if (controle.isSuperAdmin || controle.regionais.includes(reg)) {
        lojas.push({
          id: fId,
          nome: dadosLojas[j][1],
          regional: reg,
          diretoria: dadosLojas[j][3]
        });
      }
    }
  }

  // 3. Processamento de Lançamentos Próprios com Trava de 1 Visita/Loja/Dia
  var lancamentos = [];
  var moedasTotaisUser = 0;
  var moedasMesUser = 0;
  var visitasLojasDiaSet = {}; // Trava de duplicidade: 1 visita por loja por dia por usuario
  var filiaisVisitadasSet = {};
  var pontuacaoPorUsuario = {}; // [email] = { mes: X, total: Y, email: Z }

  var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (abaLanc) {
    var dLanc = abaLanc.getDataRange().getValues();
    for (var k = 1; k < dLanc.length; k++) {
      var row = dLanc[k];
      var idReg = row[0];
      var dtObj = row[1];
      var emailAutor = String(row[2] || '').toLowerCase().trim();
      var fIdLanc = ("0000" + row[3]).slice(-4);
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[5] || '').trim().toUpperCase();

      var dtStr = formatarDataSegura(dtObj);
      var dtParsed = dtObj instanceof Date ? dtObj : new Date(dtObj);
      var mReg = dtParsed.getMonth() + 1;
      var aReg = dtParsed.getFullYear();

      var ptsItem = dicionarioPremios[motivoLanc] || 1;

      if (!pontuacaoPorUsuario[emailAutor]) {
        pontuacaoPorUsuario[emailAutor] = { mes: 0, total: 0, email: emailAutor };
      }

      // Regra de Trava de Frequencia: Se houver mais de uma visita na mesma loja no mesmo dia pelo mesmo usuário, conta apenas 1 visita
      var chaveVisitaDia = emailAutor + '_' + fIdLanc + '_' + dtStr;
      var ehPrimeiraVisitaDoDia = !visitasLojasDiaSet[chaveVisitaDia];
      
      if (ehPrimeiraVisitaDoDia) {
        visitasLojasDiaSet[chaveVisitaDia] = true;
      }

      // Pontuação acumulada
      pontuacaoPorUsuario[emailAutor].total += ptsItem;
      if (mReg === mesAlvo && aReg === anoAlvo) {
        pontuacaoPorUsuario[emailAutor].mes += ptsItem;
      }

      // Métricas do usuário logado
      if (emailAutor === controle.email.toLowerCase()) {
        moedasTotaisUser += ptsItem;
        if (mReg === mesAlvo && aReg === anoAlvo) {
          moedasMesUser += ptsItem;
          if (fIdLanc && ehPrimeiraVisitaDoDia) filiaisVisitadasSet[fIdLanc] = true;
        }

        // Timeline filtrada do período alvo
        if (controle.isSuperAdmin || controle.regionais.includes(regLanc)) {
          if (mReg === mesAlvo && aReg === anoAlvo) {
            lancamentos.push({
              id: idReg,
              data: dtStr,
              filial: fIdLanc,
              lojaNome: row[4],
              motivo: row[5],
              tema: row[6] || '',
              observacao: row[7],
              evidenciaUrl: row[8] || '',
              autor: row[2]
            });
          }
        }
      }
    }
  }

  // 4. Integração de Moedas por Atendimento Social
  try {
    var ssSocial = SpreadsheetApp.openById(SPREADSHEET_SOCIAL_ID);
    
    // Aba BASE_INTERNOS
    var abaInternos = ssSocial.getSheetByName('BASE_INTERNOS');
    if (abaInternos) {
      var dInt = abaInternos.getDataRange().getValues();
      for (var x = 1; x < dInt.length; x++) {
        var dtSocial = new Date(dInt[x][0]);
        var mSoc = dtSocial.getMonth() + 1;
        var aSoc = dtSocial.getFullYear();
        var emailCoord = String(dInt[x][12] || '').toLowerCase().trim();
        var ptsSoc = dicionarioPremios['ATENDIMENTO_SOCIAL'] || 2;

        if (emailCoord) {
          if (!pontuacaoPorUsuario[emailCoord]) {
            pontuacaoPorUsuario[emailCoord] = { mes: 0, total: 0, email: emailCoord };
          }
          pontuacaoPorUsuario[emailCoord].total += ptsSoc;
          if (mSoc === mesAlvo && aSoc === anoAlvo) {
            pontuacaoPorUsuario[emailCoord].mes += ptsSoc;
            if (emailCoord === controle.email.toLowerCase()) moedasMesUser += ptsSoc;
          }
        }
      }
    }
  } catch (eSoc) {
    Logger.log('Aviso ao ler Atendimento Social: ' + eSoc.toString());
  }

  // 5. Integração de Moedas por Apurações / Feedback / Desligamentos
  try {
    var ssApur = SpreadsheetApp.openById(SPREADSHEET_APURACOES_ID);
    
    // Aba Historico_Envios (Canal = 2 moedas / Interna = 4 moedas)
    var abaEnvios = ssApur.getSheetByName('Historico_Envios');
    if (abaEnvios) {
      var dEnv = abaEnvios.getDataRange().getValues();
      for (var y = 1; y < dEnv.length; y++) {
        var dtApur = new Date(dEnv[y][0]);
        var mAp = dtApur.getMonth() + 1;
        var aAp = dtApur.getFullYear();
        var tipoApur = String(dEnv[y][1] || '').trim().toUpperCase();
        var ptsApur = (tipoApur === 'CANAL') ? 2 : 4;

        var fIdApur = ("0000" + dEnv[y][2]).slice(-4);
        var regApur = mapaLojaRegional[fIdApur] || '';

        if (regApur && controle.regionais.includes(regApur) && controle.email) {
          var targetEmail = controle.email.toLowerCase();
          if (!pontuacaoPorUsuario[targetEmail]) {
            pontuacaoPorUsuario[targetEmail] = { mes: 0, total: 0, email: targetEmail };
          }
          pontuacaoPorUsuario[targetEmail].total += ptsApur;
          if (mAp === mesAlvo && aAp === anoAlvo) {
            pontuacaoPorUsuario[targetEmail].mes += ptsApur;
            if (targetEmail === controle.email.toLowerCase()) moedasMesUser += ptsApur;
          }
        }
      }
    }
  } catch (eAp) {
    Logger.log('Aviso ao ler Apurações: ' + eAp.toString());
  }

  // 6. Placar Ranking Top 5
  var rankingArray = [];
  Object.keys(pontuacaoPorUsuario).forEach(function(em) {
    rankingArray.push(pontuacaoPorUsuario[em]);
  });
  rankingArray.sort(function(a, b) { return b.mes - a.mes; });
  var rankingTop5 = rankingArray.slice(0, 5);

  // 7. Mural de Avisos
  var avisos = [];
  var abaAvisos = ss.getSheetByName('DADOS_AVISOS');
  if (abaAvisos) {
    var dAvisos = abaAvisos.getDataRange().getValues();
    for (var a = 1; a < dAvisos.length; a++) {
      avisos.push({
        id: a,
        data: formatarDataSegura(dAvisos[a][0]),
        mensagem: dAvisos[a][1],
        autor: dAvisos[a][2]
      });
    }
  }

  // 8. Naturezas e Temas
  var naturezas = carregarNaturezasSeguras(ss);
  var temas = carregarTemasSeguras(ss);

  return {
    temAcesso: true,
    usuario: controle,
    lojas: lojas,
    lancamentos: lancamentos,
    avisos: avisos,
    naturezas: naturezas,
    temas: temas,
    gamificacao: {
      moedasTotais: moedasTotaisUser,
      moedasMes: moedasMesUser,
      metaEverest: META_EVEREST,
      percentualEverest: Math.min(100, Math.round((moedasMesUser / META_EVEREST) * 100)),
      filiaisVisitadas: Object.keys(filiaisVisitadasSet).length,
      ranking: rankingTop5,
      mesAlvo: mesAlvo,
      anoAlvo: anoAlvo
    }
  };
}

/**
 * Lê lista de naturezas da aba CONFIGURAÇÕES
 */
function carregarNaturezasSeguras(ss) {
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  var naturezas = [];
  if (!config) return ['Checklist de Loja', 'Roteiro Regional', 'Visita Presencial', 'Treinamento', 'Reunião Regional'];
  
  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var val = String(dados[i][0] || '').trim();
    if (val && !val.startsWith('TEMA_')) naturezas.push(val);
  }
  return naturezas.length ? naturezas : ['Checklist de Loja', 'Roteiro Regional', 'Visita Presencial', 'Treinamento', 'Reunião Regional'];
}

/**
 * Lê lista de temas cadastrados categorizados
 */
function carregarTemasSeguras(ss) {
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  var temas = { reuniao: [], treinamento: [] };
  if (!config) return temas;

  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var cat = String(dados[i][2] || '').trim().toUpperCase();
    var val = String(dados[i][3] || '').trim();
    if (cat === 'REUNIAO_REGIONAL' && val) temas.reuniao.push(val);
    if (cat === 'TREINAMENTO' && val) temas.treinamento.push(val);
  }
  return temas;
}

/**
 * Busca indicadores com Cache para formulário
 */
function buscarIndicadoresLoja(filialId) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'IND_LOJA_' + filialId;
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.openById(SPREADSHEET_DASH);
  var res = { vendas: '100%', nps: '85', bancoHoras: '0h', txDesligamento: '2.1%' };

  try {
    cache.put(cacheKey, JSON.stringify(res), 300);
  } catch (e) {}
  
  return res;
}
