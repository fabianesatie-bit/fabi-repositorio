/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Converte com segurança valores numéricos vindos da planilha com suporte a vírgulas (PT-BR)
 * @param {any} val - Valor bruto vindo da célula
 * @return {number} Número convertido
 */
function parseNumPtBr(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = String(val).replace(/\./g, '').replace(',', '.').trim();
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

/**
 * Extrai dia, mês e ano de uma data ou string sem problemas de fuso horário
 * @param {Date|string} dataValor - Data a ser processada
 * @return {Object} Objeto contendo {dia, mes, ano}
 */
function extrairMesAnoData(dataValor) {
  var agora = new Date();
  var res = { dia: agora.getDate(), mes: agora.getMonth() + 1, ano: agora.getFullYear() };
  if (!dataValor) return res;

  try {
    if (dataValor instanceof Date) {
      res.dia = dataValor.getDate();
      res.mes = dataValor.getMonth() + 1;
      res.ano = dataValor.getFullYear();
      return res;
    }

    var str = String(dataValor).trim();

    // Formato ISO YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      var partesIso = str.split('T')[0].split('-');
      res.ano = parseInt(partesIso[0], 10);
      res.mes = parseInt(partesIso[1], 10);
      res.dia = parseInt(partesIso[2], 10);
      return res;
    }

    // Formato DD/MM/AAAA
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
      var pBarra = str.split('/');
      res.dia = parseInt(pBarra[0], 10);
      res.mes = parseInt(pBarra[1], 10);
      res.ano = parseInt(pBarra[2].substring(0, 4), 10);
      return res;
    }

    var dObj = new Date(dataValor);
    if (!isNaN(dObj.getTime())) {
      res.dia = dObj.getDate();
      res.mes = dObj.getMonth() + 1;
      res.ano = dObj.getFullYear();
      return res;
    }
  } catch (e) {}

  return res;
}

/**
 * Carrega a aba DADOS_USUARIOS e mapeia Nome Completo e Foto do Drive indexados pelo e-mail
 * @param {Spreadsheet} ss - Instância da planilha ativa
 * @return {Object} Mapa de usuários indexado por e-mail em minúsculas
 */
function obterMapaUsuarios(ss) {
  var mapa = {};
  var aba = ss.getSheetByName('DADOS_USUARIOS');
  if (!aba) return mapa;

  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var email = String(dados[i][0] || '').toLowerCase().trim();
    if (!email) continue;

    var rawFoto = String(dados[i][6] || '').trim();
    var driveId = extrairIdDrive(rawFoto);
    var fotoUrl = driveId ? ('https://lh3.googleusercontent.com/d/' + driveId + '=w400') : '';

    mapa[email] = {
      nome: String(dados[i][1] || email).trim(),
      cargo: String(dados[i][2] || ''),
      diretoria: String(dados[i][3] || ''),
      regionais: String(dados[i][4] || ''),
      nivelAcesso: String(dados[i][5] || ''),
      foto: fotoUrl
    };
  }
  return mapa;
}

/**
 * Carrega a aba DADOS_INDICADORES e mapeia os 6 indicadores por Filial_ID (com dupla chave)
 * @param {Spreadsheet} ss - Instância da planilha ativa
 * @return {Object} Mapa de indicadores indexado pelo Filial_ID
 */
function obterMapaIndicadoresLojas(ss) {
  var mapa = {};
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
      vendas: parseNumPtBr(dados[i][3]),                  // Col D: Vendas %
      bh: parseNumPtBr(dados[i][4]),                      // Col E: BH (Banco de Horas)
      nps: parseNumPtBr(dados[i][5]),                     // Col F: NPS
      txDesligamento: parseNumPtBr(dados[i][6]),          // Col G: Tx Desl %
      coletaLixo: Math.round(parseNumPtBr(dados[i][7])),  // Col H: COLETA LIXO
      pcd: Math.round(parseNumPtBr(dados[i][8]))          // Col I: PCD (Gap)
    };

    mapa[fId] = item;
    mapa[rawNumStr] = item;
  }
  return mapa;
}

/**
 * Carrega todos os dados iniciais do Portal GP 360 com validação rápida via abas consolidadas
 * @param {number|string} filtroMes - Mês selecionado no topo do portal
 * @param {number|string} filtroAno - Ano selecionado no topo do portal
 * @return {Object} Payload completo de inicialização para a UI
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
  var mapaIndicadores = obterMapaIndicadoresLojas(ss);

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

  var chavesValidadasEspecialista = {};

  // 1. Validação via DADOS_SOCIAL
  var abaSocCons = ss.getSheetByName('DADOS_SOCIAL');
  if (abaSocCons) {
    var dSoc = abaSocCons.getDataRange().getValues();
    for (var s = 1; s < dSoc.length; s++) {
      var regSoc = String(dSoc[s][0] || '').trim().toUpperCase();
      var mesSocStr = String(dSoc[s][1] || '').trim(); // Formato MM/AAAA
      if (regSoc && mesSocStr) {
        chavesValidadasEspecialista['SOCIAL_' + regSoc + '_' + mesSocStr] = true;
      }
    }
  }

  // 2. Validação via DADOS_APUR
  var abaApurCons = ss.getSheetByName('DADOS_APUR');
  if (abaApurCons) {
    var dApur = abaApurCons.getDataRange().getValues();
    for (var a = 1; a < dApur.length; a++) {
      var regApur = String(dApur[a][0] || '').trim().toUpperCase();
      var mesApurStr = String(dApur[a][1] || '').trim(); // Formato MM/AAAA
      var classifApur = String(dApur[a][2] || '').trim().toUpperCase();

      if (regApur && mesApurStr) {
        chavesValidadasEspecialista['APUR_' + regApur + '_' + mesApurStr] = classifApur || 'VALIDADO';
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
    for (var k = 1; k < dLanc.length; k++) {
      var row = dLanc[k];
      var idReg = row[0];
      var emailAutor = String(row[2] || '').toLowerCase().trim();
      if (!emailAutor) continue;

      var rawFilialLanc = row[3];
      var numFilialLanc = normalizarFilialId(rawFilialLanc);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : String(rawFilialLanc || '');
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[4] || '').trim();
      var motivoLancUpper = motivoLanc.toUpperCase();
      var custoTotalItem = parseFloat(row[8]) || parseFloat(row[21]) || 0;

      var dataBrutaColunaM = row[12] || row[1];
      var infoDt = extrairMesAnoData(dataBrutaColunaM);
      var mReg = infoDt.mes;
      var aReg = infoDt.ano;
      var dtStr = ("0" + infoDt.dia).slice(-2) + '/' + ("0" + infoDt.mes).slice(-2) + '/' + infoDt.ano;

      var statusSalvoCol28 = String(row[27] || '').trim().toUpperCase();
      var motivoLower = motivoLanc.toLowerCase();
      var ehEspecialista = motivoLower.includes('atendimento social') || 
                           motivoLower.includes('apuraç') || 
                           motivoLower.includes('apurac') || 
                           motivoLower.includes('feedback') || 
                           motivoLower.includes('acompanhamento');

      var ehMesAtual = (mReg === mesAlvo && aReg === anoAlvo);
      var statusFinal = statusSalvoCol28;

      if (!statusFinal) {
        statusFinal = (!ehMesAtual) ? 'VALIDADO' : (ehEspecialista ? 'PENDENTE' : 'VALIDADO');
      }

      if (statusFinal === 'PENDENTE' && ehMesAtual) {
        var encSoc = !!chavesValidadasEspecialista['SOCIAL_' + regLanc + '_' + mesAnoStrRef];
        var encApur = chavesValidadasEspecialista['APUR_' + regLanc + '_' + mesAnoStrRef];

        if (encSoc || encApur) {
          statusFinal = 'VALIDADO';
          try { abaLanc.getRange(k + 1, 28).setValue('VALIDADO'); } catch (eUpd) {}
        }
      }

      var estaValidadoEspecialista = (statusFinal === 'VALIDADO');

      var ehKmAvulso = motivoLower.includes('km avulso') || 
                       motivoLower.includes('deslocamento avulso') || 
                       String(fIdLanc).toUpperCase().includes('AVULSO');

      var ptsItem = 0;
      if (!ehKmAvulso) {
        if (ehEspecialista) {
          var classifApurEspec = chavesValidadasEspecialista['APUR_' + regLanc + '_' + mesAnoStrRef] || '';
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

        if (controle.isSuperAdmin || controle.regionais.includes(regLanc) || !fIdLanc) {
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
  }

  var rankingArray = [];
  Object.keys(pontuacaoPorUsuario).forEach(function(em) {
    if (pontuacaoPorUsuario[em].mes > 0) {
      rankingArray.push(pontuacaoPorUsuario[em]);
    }
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
    visitasInLocoMes: Object.keys(filiaisVisitadasSet).length,
    reembolsoEstimadoMes: reembolsoEstimadoMes,
    distribuicaoAtividades: distribuicaoAtividadesMap,
    indicadoresLojas: mapaIndicadores,
    lancamentos: lancamentos,
    naturezas: naturezas,
    temas: temas,
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

  naturezas.sort(function(a, b) {
    return a.localeCompare(b, 'pt-BR');
  });

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
