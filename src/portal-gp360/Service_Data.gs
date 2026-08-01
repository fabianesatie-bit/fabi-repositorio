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

  var chavesValidadasSocial = {};
  try {
    var ssSoc = SpreadsheetApp.openById(SPREADSHEET_SOCIAL_ID);
    if (ssSoc) {
      var abaSocReg = ssSoc.getSheetByName('BASE_REGISTRO');
      if (abaSocReg) {
        var dSocReg = abaSocReg.getDataRange().getValues();
        for (var sr = 1; sr < dSocReg.length; sr++) {
          var fSoc = ("0000" + normalizarFilialId(dSocReg[sr][1])).slice(-4);
          var dtSoc = formatarDataSegura(dSocReg[sr][0] || new Date());
          chavesValidadasSocial['SOCIAL_' + fSoc + '_' + dtSoc] = true;
        }
      }
    }
  } catch (eSoc) {}

  var chavesValidadasApuracoes = {};
  try {
    var ssApur = SpreadsheetApp.openById(SPREADSHEET_APURACOES_ID);
    if (ssApur) {
      var abaApurHist = ssApur.getSheetByName('Historico_Envios');
      if (abaApurHist) {
        var dApurHist = abaApurHist.getDataRange().getValues();
        for (var ah = 1; ah < dApurHist.length; ah++) {
          var fApur = ("0000" + normalizarFilialId(dApurHist[ah][0])).slice(-4);
          var dtApur = formatarDataSegura(dApurHist[ah][1] || new Date());
          chavesValidadasApuracoes['APURACAO_' + fApur + '_' + dtApur] = true;
        }
      }
    }
  } catch (eApur) {}

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
      var idReg = row[0];                                         // Col A: ID_Lancamento
      var dtObj = row[1];                                         // Col B: Data_Hora
      var emailAutor = String(row[2] || '').toLowerCase().trim(); // Col C: Coordenador_Email
      var rawFilialLanc = row[3];                                // Col D: Destino_Filial_Regional
      var numFilialLanc = normalizarFilialId(rawFilialLanc);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : String(rawFilialLanc || '');
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[4] || '').trim();               // Col E: Motivo_Meta
      var motivoLancUpper = motivoLanc.toUpperCase();
      var custoTotalItem = parseFloat(row[8]) || parseFloat(row[21]) || 0; // Col I ou Col V

      var dtStr = formatarDataSegura(dtObj || row[12]);
      var dtParsed = dtObj instanceof Date ? dtObj : new Date(dtObj || row[12]);
      var mReg = dtParsed.getMonth() + 1;
      var aReg = dtParsed.getFullYear();

      // Coluna 28 (AB): Status_Validacao_Especialista
      var statusSalvoCol28 = String(row[27] || '').trim().toUpperCase();
      var motivoLower = motivoLanc.toLowerCase();
      var ehEspecialista = motivoLower.includes('atendimento social') || 
                           motivoLower.includes('apuraç') || 
                           motivoLower.includes('apurac') || 
                           motivoLower.includes('feedback') || 
                           motivoLower.includes('acompanhamento');

      var statusFinal = statusSalvoCol28;
      if (!statusFinal) {
        statusFinal = ehEspecialista ? 'PENDENTE' : 'VALIDADO';
      }

      // Reconciliação automática: Se PENDENTE, verifica se já consta na planilha externa
      if (statusFinal === 'PENDENTE') {
        var encSoc = !!chavesValidadasSocial['SOCIAL_' + fIdLanc + '_' + dtStr];
        var encApur = !!chavesValidadasApuracoes['APURACAO_' + fIdLanc + '_' + dtStr];
        if (encSoc || encApur) {
          statusFinal = 'VALIDADO';
          // Atualiza a célula fisicamente na Coluna 28 (AB) da aba DADOS_LANCAMENTOS
          try {
            abaLanc.getRange(k + 1, 28).setValue('VALIDADO');
          } catch (eUpd) {}
        }
      }

      var estaValidadoEspecialista = (statusFinal === 'VALIDADO');
      var ptsItem = parseFloat(row[9]) || dicionarioPremios[motivoLancUpper] || 1;

      if (!pontuacaoPorUsuario[emailAutor]) {
        pontuacaoPorUsuario[emailAutor] = { mes: 0, total: 0, email: emailAutor, foto: '' };
      }

      // Concede moedas apenas se o registro estiver VALIDADO
      var chaveVisitaDia = emailAutor + '_' + fIdLanc + '_' + dtStr;
      var ehPrimeiraVisitaDoDia = !visitasLojasDiaSet[chaveVisitaDia];
      
      if (ehPrimeiraVisitaDoDia && estaValidadoEspecialista) {
        visitasLojasDiaSet[chaveVisitaDia] = true;
        pontuacaoPorUsuario[emailAutor].total += ptsItem;
        if (mReg === mesAlvo && aReg === anoAlvo) {
          pontuacaoPorUsuario[emailAutor].mes += ptsItem;
        }
      }

      if (emailAutor === controle.email.toLowerCase()) {
        if (ehPrimeiraVisitaDoDia && estaValidadoEspecialista) {
          moedasTotaisUser += ptsItem;
        }

        if (mReg === mesAlvo && aReg === anoAlvo) {
          if (ehPrimeiraVisitaDoDia && estaValidadoEspecialista) {
            moedasMesUser += ptsItem;
            visitasInLocoMes++;
            if (fIdLanc) filiaisVisitadasSet[fIdLanc] = true;
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
