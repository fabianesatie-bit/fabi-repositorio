/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Extrai dia, mês e ano de objetos Date ou Strings sem problemas de fuso horário
 * @param {Date|string} dataVal - Valor de data bruto
 * @return {Object} Objeto contendo dia, mes e ano numéricos
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

    // Formato ISO YYYY-MM-DD
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

/**
 * Carrega todos os dados iniciais do Portal GP 360 e executa a conciliação ultra-rápida usando abas consolidadas
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
  var mesAnoStr = ("0" + mesAlvo).slice(-2) + '/' + anoAlvo;

  // 1. Dicionário de Usuários (Nome da Coluna B e Foto do Drive da Coluna G)
  var mapaUsuarios = {};
  var abaUsers = ss.getSheetByName('DADOS_USUARIOS');
  if (abaUsers) {
    var dUsers = abaUsers.getDataRange().getValues();
    for (var u = 1; u < dUsers.length; u++) {
      var emailU = String(dUsers[u][0] || '').toLowerCase().trim();
      if (!emailU) continue;
      
      var nomeU = String(dUsers[u][1] || '').trim(); // Coluna B: Nome
      var rawFoto = String(dUsers[u][6] || '').trim(); // Coluna G: Foto ID
      var driveId = extrairIdDrive(rawFoto);
      var fotoUrl = driveId ? ('https://lh3.googleusercontent.com/d/' + driveId + '=w400') : '';

      mapaUsuarios[emailU] = {
        nome: nomeU || emailU,
        fotoUrl: fotoUrl,
        regionalStr: String(dUsers[u][4] || '').toUpperCase()
      };
    }
  }

  // 2. Leitura da Aba DADOS_INDICADORES (Painel de Bordo Rápido)
  var painelBordo = {
    vendas: '0,00%',
    nps: '0,0',
    bh: '0,00 h',
    txDesligamento: '0,00%',
    lixo: 0,
    pcd: 0
  };

  var abaInd = ss.getSheetByName('DADOS_INDICADORES');
  if (abaInd) {
    var dInd = abaInd.getDataRange().getValues();
    var somaVendas = 0, somaBH = 0, somaNPS = 0, somaTxDesl = 0, somaLixo = 0, somaPCD = 0;
    var qtdInd = 0;

    for (var i = 1; i < dInd.length; i++) {
      var regInd = String(dInd[i][1] || '').trim().toUpperCase(); // Coluna B: Regional
      
      // Filtrar por escopo de regional do colaborador (ou todas se for Admin)
      if (controle.isSuperAdmin || controle.regionais.length === 0 || controle.regionais.includes(regInd)) {
        somaVendas += parseFloat(String(dInd[i][3]).replace(',', '.')) || 0; // Col D: Vendas
        somaBH += parseFloat(String(dInd[i][4]).replace(',', '.')) || 0;     // Col E: BH
        somaNPS += parseFloat(String(dInd[i][5]).replace(',', '.')) || 0;    // Col F: NPS
        somaTxDesl += parseFloat(String(dInd[i][6]).replace(',', '.')) || 0; // Col G: Tx Desl
        somaLixo += parseInt(dInd[i][7], 10) || 0;                          // Col H: COLETA LIXO
        somaPCD += parseInt(dInd[i][8], 10) || 0;                           // Col I: PCD
        qtdInd++;
      }
    }

    if (qtdInd > 0) {
      painelBordo.vendas = (somaVendas / qtdInd).toFixed(2).replace('.', ',') + '%';
      painelBordo.nps = (somaNPS / qtdInd).toFixed(1).replace('.', ',');
      painelBordo.bh = (somaBH / qtdInd).toFixed(2).replace('.', ',') + ' h';
      painelBordo.txDesligamento = (somaTxDesl / qtdInd).toFixed(2).replace('.', ',') + '%';
      painelBordo.lixo = somaLixo;
      painelBordo.pcd = somaPCD;
    }
  }

  // 3. Mapeamento das Lojas e Regionais
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

  // 4. Conciliação com Abas DADOS_SOCIAL e DADOS_APUR (Validação de Moedas Rápida)
  var regionaisComSocialValidadas = {};
  var abaSocial = ss.getSheetByName('DADOS_SOCIAL');
  if (abaSocial) {
    var dSoc = abaSocial.getDataRange().getValues();
    for (var s = 1; s < dSoc.length; s++) {
      var rSoc = String(dSoc[s][0] || '').trim().toUpperCase();
      var mSoc = String(dSoc[s][1] || '').trim();
      var qtdSoc = parseInt(dSoc[s][3], 10) || 0;
      if (mSoc === mesAnoStr && qtdSoc > 0) {
        regionaisComSocialValidadas[rSoc] = true;
      }
    }
  }

  var regionaisComApurValidadas = {};
  var abaApur = ss.getSheetByName('DADOS_APUR');
  if (abaApur) {
    var dApur = abaApur.getDataRange().getValues();
    for (var a = 1; a < dApur.length; a++) {
      var rApur = String(dApur[a][0] || '').trim().toUpperCase();
      var mApur = String(dApur[a][1] || '').trim();
      var clasApur = String(dApur[a][2] || '').trim().toUpperCase();
      var qtdApur = parseInt(dApur[a][3], 10) || 0;

      if (mApur === mesAnoStr && qtdApur > 0) {
        var ptsApur = 1;
        if (clasApur === 'INTERNA') ptsApur = 4;
        else if (clasApur === 'CANAL') ptsApur = 2;

        if (!regionaisComApurValidadas[rApur]) regionaisComApurValidadas[rApur] = 0;
        regionaisComApurValidadas[rApur] += ptsApur;
      }
    }
  }

  // 5. Processamento das Atividades e Cálculo de Moedas do Mês
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
      var idReg = row[0];
      var emailAutor = String(row[2] || '').toLowerCase().trim();
      if (!emailAutor) continue;

      var rawFilialLanc = row[3];
      var numFilialLanc = normalizarFilialId(rawFilialLanc);
      var fIdLanc = numFilialLanc ? ("0000" + numFilialLanc).slice(-4) : String(rawFilialLanc || '');
      var regLanc = mapaLojaRegional[fIdLanc] || '';
      var motivoLanc = String(row[4] || '').trim();
      var custoTotalItem = parseFloat(row[8]) || parseFloat(row[21]) || 0;

      var dataBrutaColM = row[12] || row[1];
      var infoDt = extrairMesAnoData(dataBrutaColM);
      var mReg = infoDt.mes;
      var aReg = infoDt.ano;
      var dtStr = ("0" + infoDt.dia).slice(-2) + '/' + ("0" + infoDt.mes).slice(-2) + '/' + infoDt.ano;

      var ehMesAtual = (mReg === mesAlvo && aReg === anoAlvo);
      var statusSalvoCol28 = String(row[27] || '').trim().toUpperCase();
      var motivoLower = motivoLanc.toLowerCase();
      var ehEspecialista = motivoLower.includes('atendimento social') || 
                           motivoLower.includes('apuraç') || 
                           motivoLower.includes('apurac') || 
                           motivoLower.includes('feedback') || 
                           motivoLower.includes('acompanhamento');

      var statusFinal = statusSalvoCol28;
      if (!statusFinal) {
        if (!ehMesAtual) {
          statusFinal = 'VALIDADO';
        } else {
          statusFinal = ehEspecialista ? 'PENDENTE' : 'VALIDADO';
        }
      }

      if (statusFinal === 'PENDENTE' && ehMesAtual) {
        var temSocial = !!regionaisComSocialValidadas[regLanc];
        var temApur = !!regionaisComApurValidadas[regLanc];

        if (temSocial || temApur) {
          statusFinal = 'VALIDADO';
          try {
            abaLanc.getRange(k + 1, 28).setValue('VALIDADO');
          } catch (eUpd) {}
        }
      }

      var estaValidadoEspecialista = (statusFinal === 'VALIDADO');
      var ehKmAvulso = motivoLower.includes('km avulso') || 
                       motivoLower.includes('deslocamento avulso') || 
                       String(fIdLanc).toUpperCase().includes('AVULSO');

      var ptsItem = 0;
      if (!ehKmAvulso) {
        if (motivoLower.includes('apuraç') || motivoLower.includes('apurac')) {
          ptsItem = regionaisComApurValidadas[regLanc] || parseFloat(row[9]) || 1;
        } else {
          ptsItem = parseFloat(row[9]) || 1;
        }
      }

      if (!pontuacaoPorUsuario[emailAutor]) {
        var uInfo = mapaUsuarios[emailAutor] || { nome: emailAutor, fotoUrl: '' };
        pontuacaoPorUsuario[emailAutor] = {
          mes: 0,
          email: emailAutor,
          nome: uInfo.nome,
          foto: uInfo.fotoUrl
        };
      }

      if (estaValidadoEspecialista && ehMesAtual) {
        pontuacaoPorUsuario[emailAutor].mes += ptsItem;
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

  // 6. Ranking Top 5 com Nome e Foto Real
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
    painelBordo: painelBordo,
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
