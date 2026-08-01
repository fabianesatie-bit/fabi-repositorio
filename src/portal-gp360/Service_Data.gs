/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

function obterDadosIniciais(filtroMes, filtroAno) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) {
    return { temAcesso: false, motivo: controle.motivo };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  var agora = new Date();
  var mesAlvo = (filtroMes !== undefined && filtroMes !== null && filtroMes !== '') ? parseInt(filtroMes) : agora.getMonth() + 1;
  var anoAlvo = (filtroAno !== undefined && filtroAno !== null && filtroAno !== '') ? parseInt(filtroAno) : agora.getFullYear();

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

  if (!dicionarioPremios['CANAL']) dicionarioPremios['CANAL'] = 2;
  if (!dicionarioPremios['INTERNA']) dicionarioPremios['INTERNA'] = 4;

  var lojas = [];
  var mapaLojaRegional = {};
  var filiaisUnicasContadas = {};

  var abaLojas = ss.getSheetByName('DADOS_LOJAS');
  if (abaLojas) {
    var dadosLojas = abaLojas.getDataRange().getValues();
    for (var j = 1; j < dadosLojas.length; j++) {
      var rawId = dadosLojas[j][0];
      var numFilial = normalizarFilialId(rawId);
      if (!numFilial) continue;

      var fId = ("0000" + numFilial).slice(-4);
      var reg = String(dadosLojas[j][2] || '').trim().toUpperCase();
      mapaLojaRegional[fId] = reg;

      if (!filiaisUnicasContadas[fId]) {
        if (controle.isSuperAdmin || controle.regionais.includes(reg)) {
          filiaisUnicasContadas[fId] = true;
          lojas.push({
            id: fId,
            nome: dadosLojas[j][1],
            regional: reg,
            diretoria: dadosLojas[j][3]
          });
        }
      }
    }
  }

  var lancamentos = [];
  var moedasTotaisUser = 0;
  var moedasMesUser = 0;
  var visitasInLocoMes = 0;
  var reembolsoEstimadoMes = 0;
  var visitasLojasDiaSet = {};
  var filiaisVisitadasSet = {};
  var pontuacaoPorUsuario = {};
  var distribuicaoAtividadesMap = {};

  var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (abaLanc) {
    var dLanc = abaLanc.getDataRange().getValues();
    for (var k = 1; k < dLanc.length; k++) {
      var row = dLanc[k];
      var idReg = row[0];
      var dtObj = row[1];
      var emailAutor = String(row[2] || '').toLowerCase().trim();
      var numFilialLanc = normalizarFilialId(row[3]);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : '';
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[5] || '').trim();
      var motivoLancUpper = motivoLanc.toUpperCase();
      var custoTotalItem = parseFloat(row[14]) || 0;

      var dtStr = formatarDataSegura(dtObj);
      var dtParsed = dtObj instanceof Date ? dtObj : new Date(dtObj);
      var mReg = dtParsed.getMonth() + 1;
      var aReg = dtParsed.getFullYear();

      var ptsItem = dicionarioPremios[motivoLancUpper] || 1;

      if (!pontuacaoPorUsuario[emailAutor]) {
        pontuacaoPorUsuario[emailAutor] = { mes: 0, total: 0, email: emailAutor, foto: '' };
      }

      var chaveVisitaDia = emailAutor + '_' + fIdLanc + '_' + dtStr;
      var ehPrimeiraVisitaDoDia = !visitasLojasDiaSet[chaveVisitaDia];
      
      if (ehPrimeiraVisitaDoDia) {
        visitasLojasDiaSet[chaveVisitaDia] = true;
        pontuacaoPorUsuario[emailAutor].total += ptsItem;
        if (mReg === mesAlvo && aReg === anoAlvo) {
          pontuacaoPorUsuario[emailAutor].mes += ptsItem;
        }
      }

      if (emailAutor === controle.email.toLowerCase()) {
        if (ehPrimeiraVisitaDoDia) {
          moedasTotaisUser += ptsItem;
        }

        if (mReg === mesAlvo && aReg === anoAlvo) {
          if (ehPrimeiraVisitaDoDia) {
            moedasMesUser += ptsItem;
            visitasInLocoMes++;
            if (fIdLanc) filiaisVisitadasSet[fIdLanc] = true;
          }
          reembolsoEstimadoMes += custoTotalItem;

          distribuicaoAtividadesMap[motivoLanc] = (distribuicaoAtividadesMap[motivoLanc] || 0) + 1;

          if (controle.isSuperAdmin || controle.regionais.includes(regLanc)) {
            lancamentos.push({
              id: idReg,
              data: dtStr,
              filial: fIdLanc,
              lojaNome: row[4],
              motivo: row[5],
              tema: row[6] || '',
              observacao: row[7],
              evidenciaUrl: row[8] || '',
              kmPercorrido: row[9] || 0,
              valorPedagio: row[10] || 0,
              valorAlimentacao: row[11] || 0,
              valorHospedagem: row[12] || 0,
              outrosCustos: row[13] || 0,
              custoTotal: custoTotalItem,
              autor: row[2]
            });
          }
        }
      }
    }
  }

  var rankingArray = [];
  Object.keys(pontuacaoPorUsuario).forEach(function(em) {
    rankingArray.push(pontuacaoPorUsuario[em]);
  });
  rankingArray.sort(function(a, b) { return b.mes - a.mes; });
  var rankingTop5 = rankingArray.slice(0, 5);

  var naturezas = carregarNaturezasSeguras(ss);
  var temas = carregarTemasSeguras(ss);

  var pctEverest = Math.min(100, Math.round((moedasMesUser / META_EVEREST) * 100));
  var faseNome = "⛺ Fase 1: Acampamento Base";
  if (pctEverest >= 100) faseNome = "🚩 ⚡ Fase 4: Bandeira no Everest!";
  else if (pctEverest >= 75) faseNome = "🏔️ Fase 3: Cume Alcançado";
  else if (pctEverest >= 40) faseNome = "🧗 Fase 2: Subida da Montanha";

  return {
    temAcesso: true,
    usuario: controle,
    lojas: lojas,
    lojasCarteiraTotal: lojas.length,
    visitasInLocoMes: visitasInLocoMes,
    reembolsoEstimadoMes: reembolsoEstimadoMes,
    distribuicaoAtividades: distribuicaoAtividadesMap,
    lancamentos: lancamentos,
    naturezas: naturezas,
    temas: temas,
    gamificacao: {
      moedasTotais: moedasTotaisUser,
      moedasMes: moedasMesUser,
      metaEverest: META_EVEREST,
      percentualEverest: pctEverest,
      faseNome: faseNome,
      filiaisVisitadas: Object.keys(filiaisVisitadasSet).length,
      ranking: rankingTop5,
      mesAlvo: mesAlvo,
      anoAlvo: anoAlvo
    }
  };
}

function carregarNaturezasSeguras(ss) {
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  var naturezas = [];
  if (!config) return ['Checklist de Loja', 'Roteiro Regional', 'Visita Presencial', 'Treinamento', 'Reunião Regional', 'Atendimento Social', 'Apurações e Feedback'];
  
  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var val = String(dados[i][0] || '').trim();
    if (val && !val.startsWith('TEMA_')) naturezas.push(val);
  }
  return naturezas.length ? naturezas : ['Checklist de Loja', 'Roteiro Regional', 'Visita Presencial', 'Treinamento', 'Reunião Regional', 'Atendimento Social', 'Apurações e Feedback'];
}

function carregarTemasSeguras(ss) {
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  var temas = { reuniao: [], treinamento: [] };
  if (!config) {
    return {
      reuniao: ['NPS', 'GMD', 'Banco de Horas', 'PCD', 'Quadro', 'Lixo Eletrônico', 'Campanha Sazonais', 'Agente Integrador'],
      treinamento: ['Liderança', 'Assédio', 'Jurídico', 'Auditoria', 'Integração', 'Atendimento 10 estrelas', 'Motivacional', 'Feedback', 'Inegociáveis', 'Apontamento']
    };
  }

  var dados = config.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var cat = String(dados[i][2] || '').trim().toUpperCase();
    var val = String(dados[i][3] || '').trim();
    if (cat === 'REUNIAO_REGIONAL' && val) temas.reuniao.push(val);
    if (cat === 'TREINAMENTO' && val) temas.treinamento.push(val);
  }

  if (temas.reuniao.length === 0) temas.reuniao = ['NPS', 'GMD', 'Banco de Horas', 'PCD', 'Quadro', 'Lixo Eletrônico', 'Campanha Sazonais', 'Agente Integrador'];
  if (temas.treinamento.length === 0) temas.treinamento = ['Liderança', 'Assédio', 'Jurídico', 'Auditoria', 'Integração', 'Atendimento 10 estrelas', 'Motivacional', 'Feedback', 'Inegociáveis', 'Apontamento'];

  return temas;
}
