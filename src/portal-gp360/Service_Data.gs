/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 * Otimizado para abertura sub-3s com validação precisa de Apurações/Social
 */

/**
 * Normalização de Filiais (Série > 3000)
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Extrai ID de um link do Google Drive
 */
function extrairIdDrive(url) {
  if (!url) return '';
  var match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

/**
 * Normaliza e extrai dia, mês e ano com alta velocidade
 */
function extrairMesAnoData(val) {
  var d = new Date();
  if (!val) return { dia: d.getDate(), mes: d.getMonth() + 1, ano: d.getFullYear() };
  if (val instanceof Date) {
    return { dia: val.getDate(), mes: val.getMonth() + 1, ano: val.getFullYear() };
  }
  var str = String(val).trim();
  var parts = str.split('/');
  if (parts.length === 3) {
    var dia = parseInt(parts[0], 10) || 1;
    var mes = parseInt(parts[1], 10) || 1;
    var ano = parseInt(parts[2], 10) || d.getFullYear();
    return { dia: dia, mes: mes, ano: ano };
  }
  var d2 = new Date(str);
  if (!isNaN(d2.getTime())) {
    return { dia: d2.getDate(), mes: d2.getMonth() + 1, ano: d2.getFullYear() };
  }
  return { dia: d.getDate(), mes: d.getMonth() + 1, ano: d.getFullYear() };
}

/**
 * Retorna o perfil de controle de acesso do usuário ativo
 */
function obterControleAcesso() {
  if (typeof obterControleAcessoSeguro === 'function') {
    return obterControleAcessoSeguro();
  }
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  return {
    temAcesso: true,
    email: email,
    nome: email.split('@')[0],
    cargo: 'Coordenador',
    nivelAcesso: 'Acesso Liberado',
    isSuperAdmin: true,
    isAdmin: true,
    regionais: []
  };
}

/**
 * Converte com segurança valores numéricos vindos da planilha
 */
function parseNumPtBr(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = String(val).replace(/%/g, '').replace(/\./g, '').replace(',', '.').trim();
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * Carrega e mapeia a aba DADOS_USUARIOS
 */
function obterMapaUsuarios(ss) {
  var cache = CacheService.getScriptCache();
  var cachedUsers = cache.get('GP360_MAPA_USUARIOS_SLIM_V5');
  if (cachedUsers) {
    try {
      var rawUsers = JSON.parse(cachedUsers);
      var mapaResult = {};
      Object.keys(rawUsers).forEach(function(email) {
        var u = rawUsers[email];
        mapaResult[email] = { nome: u[0], cargo: u[1], regional: u[2], foto: u[3] };
      });
      return mapaResult;
    } catch (e) {
      Logger.log('Recarregando mapa de usuarios...');
    }
  }

  var mapa = {};
  var rawSlim = {};
  var abaUser = ss.getSheetByName('DADOS_USUARIOS');
  if (!abaUser) return mapa;

  var dados = abaUser.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var email = String(dados[i][0] || '').toLowerCase().trim();
    if (!email) continue;

    var rawFoto = String(dados[i][6] || '').trim();
    var driveId = extrairIdDrive(rawFoto);
    var fotoUrl = driveId ? ('https://lh3.googleusercontent.com/d/' + driveId + '=w400') : '';

    var nome = dados[i][1] || email;
    var cargo = dados[i][2] || 'Colaborador';
    var regional = dados[i][4] || '';

    mapa[email] = { nome: nome, cargo: cargo, regional: regional, foto: fotoUrl };
    rawSlim[email] = [nome, cargo, regional, fotoUrl];
  }

  try { cache.put('GP360_MAPA_USUARIOS_SLIM_V5', JSON.stringify(rawSlim), 900); } catch (e) {}
  return mapa;
}

/**
 * Carrega a aba DADOS_INDICADORES
 */
function obterMapaIndicadoresLojas(ss) {
  var cache = CacheService.getScriptCache();
  var cachedInd = cache.get('GP360_MAPA_INDICADORES_SLIM_V5');
  if (cachedInd) {
    try {
      var rawMap = JSON.parse(cachedInd);
      var mapaResult = {};
      Object.keys(rawMap).forEach(function(fId) {
        var arr = rawMap[fId];
        var item = { vendas: arr[0], bh: arr[1], nps: arr[2], txDesligamento: arr[3], coletaLixo: arr[4], pcd: arr[5] };
        mapaResult[fId] = item;
        var numOnly = String(parseInt(fId, 10));
        if (numOnly && numOnly !== 'NaN') mapaResult[numOnly] = item;
      });
      return mapaResult;
    } catch (e) {
      Logger.log('Recarregando mapa de indicadores...');
    }
  }

  var mapa = {};
  var rawSlim = {};
  var abaInd = ss.getSheetByName('DADOS_INDICADORES');
  if (!abaInd) return mapa;

  var dados = abaInd.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var rawId = dados[i][0];
    var numFilial = normalizarFilialId(rawId);
    if (!numFilial) continue;

    var fId = ("0000" + numFilial).slice(-4);
    var rawNumStr = String(numFilial);

    var item = {
      vendas: parseNumPtBr(dados[i][3]),
      bh: parseNumPtBr(dados[i][4]),
      nps: parseNumPtBr(dados[i][5]),
      txDesligamento: parseNumPtBr(dados[i][6]),
      coletaLixo: Math.round(parseNumPtBr(dados[i][7])),
      pcd: Math.round(parseNumPtBr(dados[i][8]))
    };

    mapa[fId] = item;
    mapa[rawNumStr] = item;
    rawSlim[fId] = [item.vendas, item.bh, item.nps, item.txDesligamento, item.coletaLixo, item.pcd];
  }

  try { cache.put('GP360_MAPA_INDICADORES_SLIM_V5', JSON.stringify(rawSlim), 900); } catch (e) {}
  return mapa;
}

/**
 * ETAPA 1: Retorna payload leve para renderização inicial imediata
 */
function obterDadosIniciais(filtroMes, filtroAno) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) {
    return { temAcesso: false, motivo: controle.motivo };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var agora = new Date();
  var mesAlvo = (filtroMes !== undefined && filtroMes !== null && filtroMes !== '') ? parseInt(filtroMes, 10) : agora.getMonth() + 1;
  var anoAlvo = (filtroAno !== undefined && filtroAno !== null && filtroAno !== '') ? parseInt(filtroAno, 10) : agora.getFullYear();

  var mapaUsuarios = obterMapaUsuarios(ss);

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

  if (!dicionarioPremios['INTERNA']) dicionarioPremios['INTERNA'] = 4;
  if (!dicionarioPremios['CANAL']) dicionarioPremios['CANAL'] = 2;

  var chavesValidadasEspecialista = {};

  // Validação rápida via DADOS_SOCIAL
  var abaSocCons = ss.getSheetByName('DADOS_SOCIAL');
  if (abaSocCons) {
    var dSoc = abaSocCons.getDataRange().getValues();
    for (var s = 1; s < dSoc.length; s++) {
      var rawRegSoc = dSoc[s][0];
      var numFilSoc = normalizarFilialId(rawRegSoc);
      var regSocPad = numFilSoc ? ("0000" + numFilSoc).slice(-4) : String(rawRegSoc || '').trim().toUpperCase();
      var mesSocStr = String(dSoc[s][1] || '').trim();
      if (regSocPad && mesSocStr) {
        chavesValidadasEspecialista['SOCIAL_' + regSocPad + '_' + mesSocStr] = true;
      }
    }
  }

  // Validação rápida via DADOS_APUR
  var abaApurCons = ss.getSheetByName('DADOS_APUR');
  if (abaApurCons) {
    var dApur = abaApurCons.getDataRange().getValues();
    for (var a = 1; a < dApur.length; a++) {
      var rawRegApur = dApur[a][0];
      var numFilApur = normalizarFilialId(rawRegApur);
      var regApurPad = numFilApur ? ("0000" + numFilApur).slice(-4) : String(rawRegApur || '').trim().toUpperCase();
      var mesApurStr = String(dApur[a][1] || '').trim();
      var classifApur = String(dApur[a][2] || '').trim().toUpperCase();

      if (regApurPad && mesApurStr) {
        chavesValidadasEspecialista['APUR_' + regApurPad + '_' + mesApurStr] = classifApur || 'VALIDADO';
      }
    }
  }

  var lancamentos = [];
  var moedasMesUser = 0;
  var reembolsoEstimadoMes = 0;
  var filiaisVisitadasSet = {};
  var pontuacaoPorUsuario = {};
  var distribuicaoAtividadesMap = {};

  var mesAnoStrRef = ("0" + mesAlvo).slice(-2) + '/' + anoAlvo;

  var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (abaLanc) {
    var dLanc = abaLanc.getDataRange().getValues();
    var totalRows = dLanc.length;

    for (var k = 1; k < totalRows; k++) {
      var row = dLanc[k];
      var idReg = row[0];
      var emailAutor = String(row[2] || '').toLowerCase().trim();
      if (!emailAutor) continue;

      var dataBrutaColunaM = row[12] || row[1];
      var infoDt = extrairMesAnoData(dataBrutaColunaM);
      var mReg = infoDt.mes;
      var aReg = infoDt.ano;

      // Filtro de Segurança por Ano
      if (aReg !== anoAlvo) continue;

      var dtStr = ("0" + infoDt.dia).slice(-2) + '/' + ("0" + infoDt.mes).slice(-2) + '/' + infoDt.ano;
      var rawFilialLanc = row[3];
      var numFilialLanc = normalizarFilialId(rawFilialLanc);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : String(rawFilialLanc || '');
      var motivoLanc = String(row[4] || '').trim();
      var motivoLancUpper = motivoLanc.toUpperCase();
      var custoTotalItem = parseFloat(row[8]) || parseFloat(row[21]) || 0;

      var statusSalvoCol28 = String(row[27] || '').trim().toUpperCase();
      var motivoLower = motivoLanc.toLowerCase();
      var ehEspecialista = motivoLower.includes('atendimento social') || 
                           motivoLower.includes('apuraç') || 
                           motivoLower.includes('apurac') || 
                           motivoLower.includes('feedback') || 
                           motivoLower.includes('acompanhamento');

      var ehMesAtual = (mReg === mesAlvo);
      var statusFinal = statusSalvoCol28;

      var encSoc = !!chavesValidadasEspecialista['SOCIAL_' + fIdLanc + '_' + mesAnoStrRef];
      var encApur = chavesValidadasEspecialista['APUR_' + fIdLanc + '_' + mesAnoStrRef];

      if (ehEspecialista && (encSoc || encApur)) {
        statusFinal = 'VALIDADO';
      }

      if (!statusFinal) {
        statusFinal = (!ehMesAtual) ? 'VALIDADO' : (ehEspecialista ? 'PENDENTE' : 'VALIDADO');
      }

      var estaValidadoEspecialista = (statusFinal === 'VALIDADO');

      var ehKmAvulso = motivoLower.includes('km avulso') || 
                       motivoLower.includes('deslocamento avulso') || 
                       String(fIdLanc).toUpperCase().includes('AVULSO');

      var ptsItem = 0;
      if (!ehKmAvulso) {
        if (ehEspecialista) {
          var classifApurEspec = chavesValidadasEspecialista['APUR_' + fIdLanc + '_' + mesAnoStrRef] || '';
          if (classifApurEspec === 'INTERNA') ptsItem = 4;
          else if (classifApurEspec === 'CANAL') ptsItem = 2;
          else ptsItem = dicionarioPremios[motivoLancUpper] || parseFloat(row[9]) || 1;
        } else {
          ptsItem = parseFloat(row[9]) || dicionarioPremios[motivoLancUpper] || 1;
        }
      }

      var usrInfo = mapaUsuarios[emailAutor] || { nome: emailAutor, foto: '' };

      if (!pontuacaoPorUsuario[emailAutor]) {
        pontuacaoPorUsuario[emailAutor] = { 
          mes: 0, 
          total: 0, 
          email: emailAutor, 
          nome: usrInfo.nome, 
          foto: usrInfo.foto 
        };
      }

      if (estaValidadoEspecialista && ehMesAtual) {
        pontuacaoPorUsuario[emailAutor].mes += ptsItem;
        pontuacaoPorUsuario[emailAutor].total += ptsItem;
      }

      if (emailAutor === controle.email.toLowerCase() && ehMesAtual) {
        if (estaValidadoEspecialista) {
          moedasMesUser += ptsItem;
          if (fIdLanc && !ehKmAvulso) filiaisVisitadasSet[fIdLanc] = true;
        }

        reembolsoEstimadoMes += custoTotalItem;

        if (motivoLanc) {
          distribuicaoAtividadesMap[motivoLanc] = (distribuicaoAtividadesMap[motivoLanc] || 0) + 1;
        }

        lancamentos.push({
          id: idReg,
          data: dtStr,
          filial: fIdLanc,
          motivo: motivoLanc,
          tema: row[18] || '',
          observacao: row[10] || '',
          evidenciaUrl: row[11] || '',
          kmPercorrido: row[15] || 0,
          custoKm: row[16] || 0,
          tipoRoteiro: row[17] || '',
          pessoasImpactadas: row[19] || 0,
          tempoGasto: row[20] || 0,
          valorPedagio: row[25] || 0,
          valorAlimentacao: row[22] || 0,
          valorHospedagem: row[23] || 0,
          valorAereo: row[24] || 0,
          valorEstacionamento: row[26] || 0,
          custoTotal: custoTotalItem,
          statusValidacao: statusFinal,
          validadoEspecialista: estaValidadoEspecialista,
          autor: row[2]
        });
      }
    }
  }

  var rankingArray = [];
  Object.keys(pontuacaoPorUsuario).forEach(function(em) {
    if (pontuacaoPorUsuario[em].mes > 0) {
      rankingArray.push(pontuacaoPorUsuario[em]);
    }
  });

  rankingArray.sort(function(a, b) { return b.mes - a.mes; });
  var rankingTop5 = rankingArray.slice(0, 5);

  var pctEverest = Math.min(100, Math.round((moedasMesUser / META_EVEREST) * 100));
  var faseNome = "⛺ Fase 1: Acampamento Base";
  if (pctEverest >= 100) faseNome = "🚩 ⚡ Fase 4: Bandeira no Everest!";
  else if (pctEverest >= 75) faseNome = "🏔️ Fase 3: Cume Alcançado";
  else if (pctEverest >= 40) faseNome = "🧗 Fase 2: Subida da Montanha";

  return {
    temAcesso: true,
    usuario: controle,
    visitasInLocoMes: Object.keys(filiaisVisitadasSet).length,
    reembolsoEstimadoMes: reembolsoEstimadoMes,
    distribuicaoAtividades: distribuicaoAtividadesMap,
    lancamentos: lancamentos,
    dicionarioPremios: dicionarioPremios,
    gamificacao: {
      moedasTotais: moedasMesUser,
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

/**
 * ETAPA 2: Carregamento assíncrono secundário
 */
function obterDadosComplementares() {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) return {};

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mapaIndicadores = obterMapaIndicadoresLojas(ss);

  var lojas = [];
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

  var naturezas = carregarNaturezasSeguras(ss);
  var temas = carregarTemasSeguras(ss);

  return {
    lojas: lojas,
    lojasCarteiraTotal: lojas.length,
    indicadoresLojas: mapaIndicadores,
    naturezas: naturezas,
    temas: temas
  };
}

function carregarNaturezasSeguras(ss) {
  var config = ss.getSheetByName('CONFIGURAÇÕES');
  var naturezas = [];
  if (config) {
    var dados = config.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      var val = String(dados[i][0] || '').trim();
      if (val && !val.startsWith('TEMA_')) naturezas.push(val);
    }
  }

  if (naturezas.length === 0) {
    naturezas = ['Apurações e Feedback', 'Atendimento Social', 'Checklist de Loja', 'Reunião Regional', 'Roteiro Regional', 'Treinamento', 'Visita Presencial'];
  }

  naturezas.sort(function(a, b) { return a.localeCompare(b, 'pt-BR'); });
  return naturezas;
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
