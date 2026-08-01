/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

function extrairMesAnoData(dataVal) {
  if (!dataVal) return { dia: 0, mes: 0, ano: 0 };
  try {
    if (dataVal instanceof Date) {
      return { 
        dia: dataVal.getDate(), 
        mes: dataVal.getMonth() + 1, 
        ano: dataVal.getFullYear() 
      };
    }

    var str = String(dataVal).trim();

    // Formato DD/MM/YYYY ou DD/MM/YYYY HH:mm:ss
    if (str.indexOf('/') !== -1) {
      var partesBarra = str.split(' ')[0].split('/');
      if (partesBarra.length === 3) {
        return {
          dia: parseInt(partesBarra[0], 10),
          mes: parseInt(partesBarra[1], 10),
          ano: parseInt(partesBarra[2], 10)
        };
      }
    }

    // Formato YYYY-MM-DD
    if (str.indexOf('-') !== -1) {
      var partesIso = str.split('T')[0].split('-');
      if (partesIso.length === 3 && partesIso[0].length === 4) {
        return {
          dia: parseInt(partesIso[2], 10),
          mes: parseInt(partesIso[1], 10),
          ano: parseInt(partesIso[0], 10)
        };
      }
    }

    var dObj = new Date(dataVal);
    if (!isNaN(dObj.getTime())) {
      return { 
        dia: dObj.getDate(), 
        mes: dObj.getMonth() + 1, 
        ano: dObj.getFullYear() 
      };
    }
  } catch (e) {}

  return { dia: 0, mes: 0, ano: 0 };
}

function obterDadosIniciais(filtroMes, filtroAno) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) {
    return { temAcesso: false, motivo: controle.motivo };
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var agora = new Date();
  var mesAlvo = (filtroMes !== undefined && filtroMes !== null && filtroMes !== '') ? parseInt(filtroMes, 10) : agora.getMonth() + 1;
  var anoAlvo = (filtroAno !== undefined && filtroAno !== null && filtroAno !== '') ? parseInt(filtroAno, 10) : agora.getFullYear();

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

  function registrarChavesValidadas(mapa, prefixo, filial, email, dataStr) {
    if (filial) {
      mapa[prefixo + '_' + filial] = true;
    }
    if (email) {
      mapa[prefixo + '_EMAIL_' + email] = true;
    }

    if (dataStr) {
      var infoDt = extrairMesAnoData(dataStr);
      if (infoDt.mes > 0) {
        var dtFmt = ("0" + infoDt.dia).slice(-2) + '/' + ("0" + infoDt.mes).slice(-2) + '/' + infoDt.ano;
        if (filial) {
          mapa[prefixo + '_' + filial + '_' + dtFmt] = true;
          mapa[prefixo + '_' + filial + '_M' + infoDt.mes + '_A' + infoDt.ano] = true;
        }
        if (email) {
          mapa[prefixo + '_EMAIL_' + email + '_' + dtFmt] = true;
          mapa[prefixo + '_EMAIL_' + email + '_M' + infoDt.mes + '_A' + infoDt.ano] = true;
        }
      }
    }
  }

  var chavesValidadasSocial = {};
  try {
    var ssSoc = SpreadsheetApp.openById(SPREADSHEET_SOCIAL_ID);
    if (ssSoc) {
      // 1. Aba BASE_REGISTRO (Coluna C = Index 2 é a Filial!)
      var abaSocReg = ssSoc.getSheetByName('BASE_REGISTRO');
      if (abaSocReg) {
        var dSocReg = abaSocReg.getDataRange().getValues();
        for (var sr = 1; sr < dSocReg.length; sr++) {
          var numSoc = normalizarFilialId(dSocReg[sr][2] || dSocReg[sr][1]);
          var fSoc = numSoc ? ("0000" + numSoc).slice(-4) : '';
          var dtSoc = formatarDataSegura(dSocReg[sr][0] || new Date());
          registrarChavesValidadas(chavesValidadasSocial, 'SOCIAL', fSoc, null, dtSoc);
        }
      }

      // 2. Aba BASE_INTERNOS
      var abaSocInt = ssSoc.getSheetByName('BASE_INTERNOS');
      if (abaSocInt) {
        var dSocInt = abaSocInt.getDataRange().getValues();
        for (var si = 1; si < dSocInt.length; si++) {
          var numSocInt = normalizarFilialId(dSocInt[si][1]);
          var fSocInt = numSocInt ? ("0000" + numSocInt).slice(-4) : '';
          var dtSocInt = formatarDataSegura(dSocInt[si][0] || new Date());
          var emailCoordenador = String(dSocInt[si][2] || '').toLowerCase().trim();
          registrarChavesValidadas(chavesValidadasSocial, 'SOCIAL', fSocInt, emailCoordenador, dtSocInt);
        }
      }
    }
  } catch (eSoc) {
    Logger.log('Erro ao ler planilha Social: ' + eSoc.toString());
  }

  var chavesValidadasApuracoes = {};
  try {
    var ssApur = SpreadsheetApp.openById(SPREADSHEET_APURACOES_ID);
    if (ssApur) {
      // 1. Aba Historico_Envios
      var abaApurHist = ssApur.getSheetByName('Historico_Envios');
      if (abaApurHist) {
        var dApurHist = abaApurHist.getDataRange().getValues();
        for (var ah = 1; ah < dApurHist.length; ah++) {
          var numApur = normalizarFilialId(dApurHist[ah][0]);
          var fApur = numApur ? ("0000" + numApur).slice(-4) : '';
          var dtApur = formatarDataSegura(dApurHist[ah][1] || new Date());
          registrarChavesValidadas(chavesValidadasApuracoes, 'APURACAO', fApur, null, dtApur);
        }
      }

      // 2. Aba Intervencoes_Feedback (Coluna B: Filial, Coluna G: Data - Index 6)
      var abaApurFeed = ssApur.getSheetByName('Intervencoes_Feedback');
      if (abaApurFeed) {
        var dApurFeed = abaApurFeed.getDataRange().getValues();
        for (var af = 1; af < dApurFeed.length; af++) {
          var numFeed = normalizarFilialId(dApurFeed[af][1]);
          var fFeed = numFeed ? ("0000" + numFeed).slice(-4) : '';
          var dtFeed = formatarDataSegura(dApurFeed[af][6] || dApurFeed[af][0] || new Date());
          var emailFeed = String(dApurFeed[af][7] || dApurFeed[af][8] || '').toLowerCase().trim();
          registrarChavesValidadas(chavesValidadasApuracoes, 'APURACAO', fFeed, emailFeed, dtFeed);
        }
      }

      // 3. Aba Historico_Desligamentos
      var abaApurDesl = ssApur.getSheetByName('Historico_Desligamentos');
      if (abaApurDesl) {
        var dApurDesl = abaApurDesl.getDataRange().getValues();
        for (var ad = 1; ad < dApurDesl.length; ad++) {
          var numDesl = normalizarFilialId(dApurDesl[ad][0]);
          var fDesl = numDesl ? ("0000" + numDesl).slice(-4) : '';
          var dtDesl = formatarDataSegura(dApurDesl[ad][1] || new Date());
          var emailReg = String(dApurDesl[ad][2] || '').toLowerCase().trim();
          registrarChavesValidadas(chavesValidadasApuracoes, 'APURACAO', fDesl, emailReg, dtDesl);
        }
      }
    }
  } catch (eApur) {
    Logger.log('Erro ao ler planilha Apurações: ' + eApur.toString());
  }

  var lancamentos = [];
  var moedasMesUser = 0;
  var reembolsoEstimadoMes = 0;
  var filiaisVisitadasSet = {};
  var pontuacaoPorUsuario = {};
  var distribuicaoAtividadesMap = {};

  var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (abaLanc) {
    var dLanc = abaLanc.getDataRange().getValues();
    for (var k = 1; k < dLanc.length; k++) {
      var row = dLanc[k];
      var idReg = row[0];                                         // Col A: ID_Lancamento
      var emailAutor = String(row[2] || '').toLowerCase().trim(); // Col C: Coordenador_Email
      if (!emailAutor) continue;

      var rawFilialLanc = row[3];                                // Col D: Destino_Filial_Regional
      var numFilialLanc = normalizarFilialId(rawFilialLanc);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : String(rawFilialLanc || '');
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[4] || '').trim();               // Col E: Motivo_Meta
      var motivoLancUpper = motivoLanc.toUpperCase();
      var custoTotalItem = parseFloat(row[8]) || parseFloat(row[21]) || 0; // Col I ou Col V

      // DATA PRIMÁRIA ABSOLUTA: Coluna M (index 12: Data_Ini)
      var dataBrutaColunaM = row[12] || row[1];
      var infoDt = extrairMesAnoData(dataBrutaColunaM);
      var mReg = infoDt.mes;
      var aReg = infoDt.ano;
      var dtStr = ("0" + infoDt.dia).slice(-2) + '/' + ("0" + infoDt.mes).slice(-2) + '/' + infoDt.ano;

      // Coluna 28 (Col AB): Status_Validacao_Especialista (index 27)
      var statusSalvoCol28 = String(row[27] || '').trim().toUpperCase();
      var motivoLower = motivoLanc.toLowerCase();
      var ehEspecialista = motivoLower.includes('atendimento social') || 
                           motivoLower.includes('apuraç') || 
                           motivoLower.includes('apurac') || 
                           motivoLower.includes('feedback') || 
                           motivoLower.includes('acompanhamento');

      // REGRA TEMPORAL: Para meses passados, se Col AB estiver vazia, trata como VALIDADO.
      var ehMesAtual = (mReg === mesAlvo && aReg === anoAlvo);
      var statusFinal = statusSalvoCol28;

      if (!statusFinal) {
        if (!ehMesAtual) {
          statusFinal = 'VALIDADO';
        } else {
          statusFinal = ehEspecialista ? 'PENDENTE' : 'VALIDADO';
        }
      }

      if (statusFinal === 'PENDENTE' && ehMesAtual) {
        var encSoc = !!chavesValidadasSocial['SOCIAL_' + fIdLanc] || 
                     !!chavesValidadasSocial['SOCIAL_' + fIdLanc + '_' + dtStr] ||
                     !!chavesValidadasSocial['SOCIAL_' + fIdLanc + '_M' + mReg + '_A' + aReg] ||
                     !!chavesValidadasSocial['SOCIAL_EMAIL_' + emailAutor];

        var encApur = !!chavesValidadasApuracoes['APURACAO_' + fIdLanc] ||
                      !!chavesValidadasApuracoes['APURACAO_' + fIdLanc + '_' + dtStr] ||
                      !!chavesValidadasApuracoes['APURACAO_' + fIdLanc + '_M' + mReg + '_A' + aReg] ||
                      !!chavesValidadasApuracoes['APURACAO_EMAIL_' + emailAutor];

        if (encSoc || encApur) {
          statusFinal = 'VALIDADO';
          try {
            abaLanc.getRange(k + 1, 28).setValue('VALIDADO');
          } catch (eUpd) {}
        }
      }

      var estaValidadoEspecialista = (statusFinal === 'VALIDADO');

      // REGRA DE MOEDAS: KM Avulso NAO ganha moeda
      var ehKmAvulso = motivoLower.includes('km avulso') || 
                       motivoLower.includes('deslocamento avulso') || 
                       String(fIdLanc).toUpperCase().includes('AVULSO');

      var ptsItem = 0;
      if (!ehKmAvulso) {
        if (ehMesAtual && ehEspecialista) {
          ptsItem = dicionarioPremios[motivoLancUpper] || parseFloat(row[9]) || 1;
        } else {
          ptsItem = parseFloat(row[9]) || dicionarioPremios[motivoLancUpper] || 1;
        }
      }

      if (!pontuacaoPorUsuario[emailAutor]) {
        pontuacaoPorUsuario[emailAutor] = { mes: 0, total: 0, email: emailAutor, foto: '' };
      }

      // ACUMULA MOEDAS EXCLUSIVAMENTE PARA O MÊS VIGENTE SELEÇÃO
      if (estaValidadoEspecialista && ehMesAtual) {
        pontuacaoPorUsuario[emailAutor].mes += ptsItem;
        pontuacaoPorUsuario[emailAutor].total += ptsItem; // Total Acumulado restrito ao mês vigente
      }

      if (emailAutor === controle.email.toLowerCase() && ehMesAtual) {
        if (estaValidadoEspecialista) {
          moedasMesUser += ptsItem;
          if (fIdLanc && !ehKmAvulso) {
            filiaisVisitadasSet[fIdLanc] = true;
          }
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
            tema: row[18] || '',                             // Col S: Sub temas
            observacao: row[10] || '',                       // Col K: Observacoes
            evidenciaUrl: row[11] || '',                     // Col L: Link_Evidencia
            kmPercorrido: row[15] || 0,                      // Col P: Qde_Km
            custoKm: row[16] || 0,                           // Col Q: Total Km
            tipoRoteiro: row[17] || '',                      // Col R: Tipo_Roteiro
            pessoasImpactadas: row[19] || 0,                 // Col T: Pessoas Impactas
            tempoGasto: row[20] || 0,                        // Col U: Tempo Gasto
            valorPedagio: row[25] || 0,                      // Col Z: Pedagio
            valorAlimentacao: row[22] || 0,                  // Col W: Alimentação
            valorHospedagem: row[23] || 0,                   // Col X: Hospedagem
            valorAereo: row[24] || 0,                        // Col Y: Aereo
            valorEstacionamento: row[26] || 0,               // Col AA: Estacionamento
            custoTotal: custoTotalItem,
            statusValidacao: statusFinal,                    // Col AB: Status_Validacao_Especialista
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
    lancamentos: lancamentos,
    naturezas: naturezas,
    temas: temas,
    dicionarioPremios: dicionarioPremios,
    gamificacao: {
      moedasTotais: moedasMesUser, // Total acumulado restrito ao mês vigente
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
