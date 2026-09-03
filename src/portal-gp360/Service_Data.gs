/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Data.gs
 * Subpasta Monorepo: src/portal-gp360/
 * Mapeamento e Consultas de Indicadores Operacionais
 */

function parseNumPtBr(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = String(val).replace(/%/g, '').replace(/\./g, '').replace(',', '.').trim();
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

// =============================================================================
// CARREGAMENTO DE DADOS E CONSULTAS DE INDICADORES (OTIMIZADO E RÁPIDO)
// =============================================================================

function obterDadosIniciais(mesAlvo, anoAlvo) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      return { bloqueado: true, mensagem: controle.erro };
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let usuario = { 
      email: controle.email, 
      nome: controle.nome, 
      cargo: controle.cargo, 
      infoLinha2: controle.isSuperAdmin ? "Visão Nacional" : controle.regionais.join(', '), 
      regionais: controle.regionais, 
      diretoriasAtendidas: controle.diretoriasAtendidas || [],
      isAdmin: controle.isAdmin, 
      isSuperAdmin: controle.isSuperAdmin, 
      isConfigAdmin: controle.isConfigAdmin, 
      isGerenteGP: controle.isGerenteGP,
      foto: "" 
    };

    let mapaUsuarios = {}; 
    const sheetUsuarios = ss.getSheetByName('DADOS_USUARIOS');
    let listaDiretoriasUnicas = new Set();
    if (sheetUsuarios) {
      const dadosUsuarios = sheetUsuarios.getDataRange().getValues();
      for (let i = 1; i < dadosUsuarios.length; i++) {
        const emailLinha = dadosUsuarios[i][0] ? String(dadosUsuarios[i][0]).toLowerCase().trim() : "";
        if (emailLinha) {
           let dirRaw = dadosUsuarios[i][3] ? String(dadosUsuarios[i][3]).trim() : "";
           if (dirRaw) {
              dirRaw.split(',').forEach(d => listaDiretoriasUnicas.add(d.trim()));
           }
           let fotoRaw = (dadosUsuarios[i].length > 6 && dadosUsuarios[i][6]) ? String(dadosUsuarios[i][6]).trim() : "";
           let fotoId = extrairIdDrive(fotoRaw);
           let fotoPronta = fotoId ? (fotoId.startsWith("http") ? fotoId : "https://lh3.googleusercontent.com/d/" + fotoId) : "";

           mapaUsuarios[emailLinha] = { 
               nome: dadosUsuarios[i][1] || 'GP', 
               foto: fotoPronta,
               regionais: dadosUsuarios[i][4] ? String(dadosUsuarios[i][4]).split(',').map(r => r.trim()) : []
           };
        }
        if (emailLinha === emailLogado) {
          usuario.foto = mapaUsuarios[emailLinha].foto;
        }
      }
    }
    registrarAuditoria("Login", "Acesso ao Portal Concluído");

    let placarMoedasTotal = {};
    let placarMoedasMes = {};
    let placarMoedasMesAnterior = {};
    let gastoTotalAcumulado = 0;
    let gastoMesAnterior = 0;
    let setVisitasFisicasMes = new Set();
    let setVisitasFisicasMesAnterior = new Set();

    const dHoje = new Date();
    const mesAtual = typeof mesAlvo === 'number' ? mesAlvo : dHoje.getMonth();
    const anoAtual = typeof anoAlvo === 'number' ? anoAlvo : dHoje.getFullYear();
    
    // Suporte retroativo (Mês Vigente vs Mês Anterior)
    const mesAnterior = mesAtual === 0 ? 11 : mesAtual - 1;
    const anoAnterior = mesAtual === 0 ? anoAtual - 1 : anoAtual;

    const meusLancamentos = []; 
    const sheetLancamentos = ss.getSheetByName('DADOS_LANCAMENTOS'); 

    if (sheetLancamentos && sheetLancamentos.getLastRow() > 1) {
      const dadosLancamentos = sheetLancamentos.getDataRange().getValues();
      for (let i = 1; i < dadosLancamentos.length; i++) {
        const emailAutor = dadosLancamentos[i][2] ? String(dadosLancamentos[i][2]).toLowerCase().trim() : "";
        if (emailAutor) {
           if (!placarMoedasTotal[emailAutor]) placarMoedasTotal[emailAutor] = 0;
           if (!placarMoedasMes[emailAutor]) placarMoedasMes[emailAutor] = 0;
           if (!placarMoedasMesAnterior[emailAutor]) placarMoedasMesAnterior[emailAutor] = 0;

           let valorMoedaRaw = String(dadosLancamentos[i][9] || "0").replace(',', '.').trim();
           let moedasParse = parseFloat(valorMoedaRaw);
           if (!isNaN(moedasParse)) {
               placarMoedasTotal[emailAutor] += moedasParse;

               let dtAcaoRaw = obterDataRawSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
               if (dtAcaoRaw > 0) {
                 let dtObj = new Date(dtAcaoRaw);
                 if (!isNaN(dtObj.getTime())) {
                   if (dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual) {
                     placarMoedasMes[emailAutor] += moedasParse;
                   }
                   if (dtObj.getMonth() === mesAnterior && dtObj.getFullYear() === anoAnterior) {
                     placarMoedasMesAnterior[emailAutor] += moedasParse;
                   }
                 }
               }
           }
        }
        let visivelNoHistorico = false;
        if (usuario.isSuperAdmin) {
          visivelNoHistorico = true;
        } else if (usuario.isAdmin) {
          let regAutor = mapaUsuarios[emailAutor] ? mapaUsuarios[emailAutor].regionais : [];
          let overlap = regAutor.some(r => usuario.regionais.includes(r));
          if (overlap || emailAutor === emailLogado) visivelNoHistorico = true;
        } else {
          if (emailAutor === emailLogado) visivelNoHistorico = true;
        }
        if (visivelNoHistorico) {
          let kmCusto = Number(dadosLancamentos[i][16]) || 0;
          let alimentacao = Number(dadosLancamentos[i][22]) || 0;
          let hospedagem = Number(dadosLancamentos[i][23]) || 0;
          let aereo = Number(dadosLancamentos[i][24]) || 0;
          let pedagio = Number(dadosLancamentos[i][25]) || 0;
          let estacionamento = Number(dadosLancamentos[i][26]) || 0;
          let gastoSomaReal = kmCusto + alimentacao + hospedagem + aereo + pedagio + estacionamento;

          let dtViagemReal = obterDataRawSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
          if (emailAutor === emailLogado) {
              let dtObj = new Date(dtViagemReal);
              if (!isNaN(dtObj.getTime())) {
                  let isAtual = (dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual);
                  let isAnterior = (dtObj.getMonth() === mesAnterior && dtObj.getFullYear() === anoAnterior);

                  if (isAtual) gastoTotalAcumulado += gastoSomaReal;
                  if (isAnterior) gastoMesAnterior += gastoSomaReal;

                  let rot = String(dadosLancamentos[i][17] || "").trim();
                  if (rot.includes("Visita in loco")) {
                      let dtVFormat = formatarDataSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
                      let dst = String(dadosLancamentos[i][3]).trim();
                      if (isAtual) setVisitasFisicasMes.add(emailAutor + "|" + dtVFormat + "|" + dst);
                      if (isAnterior) setVisitasFisicasMesAnterior.add(emailAutor + "|" + dtVFormat + "|" + dst);
                  }
              }
          }
          meusLancamentos.push({
            id: dadosLancamentos[i][0], 
            dataRegistro: formatarDataSegura(dadosLancamentos[i][1]), 
            autorNome: mapaUsuarios[emailAutor] ? mapaUsuarios[emailAutor].nome : emailAutor,
            isOwner: (emailAutor === emailLogado), 
            destino: dadosLancamentos[i][3], 
            motivo: dadosLancamentos[i][4], 
            gastoTotal: gastoSomaReal, 
            dataViagem: formatarDataSegura(dadosLancamentos[i][12]), 
            dataViagemRaw: dtViagemReal, 
            kmValor: Number(dadosLancamentos[i][14]) || 0, 
            kmQtd: Number(dadosLancamentos[i][15]) || 0, 
            kmCusto: kmCusto, 
            alimentacao: alimentacao,
            hospedagem: hospedagem,
            aereo: aereo,
            pedagio: pedagio,
            estacionamento: estacionamento,
            observacoes: String(dadosLancamentos[i][10] || "").trim(), 
            roteiro: String(dadosLancamentos[i][17] || "").trim(), 
            subTema: String(dadosLancamentos[i][18] || "").trim(), 
            pessoasImpactadas: Number(dadosLancamentos[i][19]) || 0, 
            tempoGasto: Number(dadosLancamentos[i][20]) || 0,
            linkEvidencia: String(dadosLancamentos[i][11] || "").trim()
          });
        }
      }
    }

    meusLancamentos.reverse(); 
    let historicoLimitado = meusLancamentos.slice(0, 400);

    let userMoedasTotal = placarMoedasTotal[emailLogado] || 0;
    let userMoedasMes = placarMoedasMes[emailLogado] || 0;
    let userMoedasMesAnterior = placarMoedasMesAnterior[emailLogado] || 0;

    // === NOVA REGRA: FASE BASEADA APENAS NO MÊS VIGENTE ===
    let faseMontanha = 1;
    if (userMoedasMes >= 120) faseMontanha = 4;
    else if (userMoedasMes >= 80) faseMontanha = 3;
    else if (userMoedasMes >= 40) faseMontanha = 2;

    let ranking = Object.keys(placarMoedasTotal).map(email => {
        let nomeExibicao = 'GP';
        let fotoExibicao = '';
        if (mapaUsuarios[email]) {
            nomeExibicao = mapaUsuarios[email].nome;
            fotoExibicao = mapaUsuarios[email].foto;
        } else {
            let partesEmail = email.split('@')[0].split('.');
            nomeExibicao = partesEmail.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
        }
        
        let totalU = placarMoedasTotal[email] || 0;
        let moedasMesU = placarMoedasMes[email] || 0;
        let moedasMesAnteriorU = placarMoedasMesAnterior[email] || 0;

        // === NOVA REGRA RANKING: FASE BASEADA APENAS NO MÊS VIGENTE ===
        let faseU = 1;
        if (moedasMesU >= 120) faseU = 4;
        else if (moedasMesU >= 80) faseU = 3;
        else if (moedasMesU >= 40) faseU = 2;

        return { 
          email: email, 
          nome: nomeExibicao, 
          foto: fotoExibicao, 
          moedas: totalU,
          moedasMes: moedasMesU,
          moedasMesAnterior: moedasMesAnteriorU,
          fase: faseU
        };
    });

    const lojasFiltradasMap = new Map();
    const sheetLojas = ss.getSheetByName('DADOS_LOJAS');
    if (sheetLojas && sheetLojas.getLastRow() > 1) {
      const dadosLojas = sheetLojas.getDataRange().getValues();
      for (let i = 1; i < dadosLojas.length; i++) {
        const regionalLoja = dadosLojas[i][2] ? String(dadosLojas[i][2]).trim() : "";
        if (usuario.isSuperAdmin || usuario.regionais.includes(regionalLoja)) {
          let idLojaClean = normalizarFilialId(dadosLojas[i][0]);
          let idLojaNum = parseInt(idLojaClean, 10);
          if (!isNaN(idLojaNum)) {
              if (!lojasFiltradasMap.has(idLojaNum)) {
                  lojasFiltradasMap.set(idLojaNum, { id: idLojaNum, nome: dadosLojas[i][1], regional: regionalLoja });
              }
          }
        }
      }
    }
    const lojasFiltradas = Array.from(lojasFiltradasMap.values());
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    const avisos = [];
    if (sheetAvisos && sheetAvisos.getLastRow() > 1) {
       const dadosAvisos = sheetAvisos.getDataRange().getValues();
       for(let i = 1; i < dadosAvisos.length; i++) {
          avisos.push({ data: formatarDataSegura(dadosAvisos[i][0]), autor: dadosAvisos[i][1], mensagem: dadosAvisos[i][2] });
       }
       avisos.reverse(); 
    }
    let naturezas = carregarNaturezasSeguras(ss);
    let diretoriasAr = Array.from(listaDiretoriasUnicas).filter(Boolean).sort();

    return { 
      bloqueado: false, usuario: usuario, lojas: lojasFiltradas, qtdLojasCarteira: lojasFiltradas.length,
      qtdVisitas: setVisitasFisicasMes.size,
      qtdVisitasAnterior: setVisitasFisicasMesAnterior.size,
      moedas: userMoedasTotal,
      moedasMes: userMoedasMes,
      moedasMesAnterior: userMoedasMesAnterior,
      faseMontanha: faseMontanha,
      gastos: gastoTotalAcumulado,
      gastosMesAnterior: gastoMesAnterior,
      historico: historicoLimitado, 
      avisos: avisos, ranking: ranking, naturezas: naturezas, diretorias: diretoriasAr,
      mesAtualNum: mesAtual,
      anoAtualNum: anoAtual,
      mesAnteriorNum: mesAnterior,
      anoAnteriorNum: anoAnterior
    };
  } catch (e) { return { erro: e.message }; }
}

/**
 * Consulta indicadores de desempenho de filiais na base DB_DASH com cache de 5 minutos.
 * Ajustado para cruzar filiais corretamente pelas Colunas certas de cada historico.
 */
function buscarIndicadoresLoja(filialId) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const verificacao = validarAcessoFilial(emailLogado, filialId);
    if (!verificacao.autorizado) {
      return { erro: "ACESSO NEGADO: Você não tem permissão para ler dados desta filial." };
    }
    let targetFilial = parseInt(normalizarFilialId(filialId), 10);
    const cache = CacheService.getScriptCache();
    const cacheKey = 'IND_LOJA_' + targetFilial;
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);

    const ss = SpreadsheetApp.openById(SPREADSHEET_DASH);
    let nps = "-", venda = "-", bh = "-", lixo = "-", quadro = "-", txdes = "-";
    const abasConsultadas = [
      'HISTORICO_VENDA_ANO', 
      'HISTORICO_NPS', 
      'HISTORICO_BH_ACUMULADO', 
      'HISTORICO_QUADRO', 
      'HISTORICO_TXDESL',
      'HISTORICO_LIXO',
      'HISTORICO_COLETAS',
      'HISTORICO_LIXO_ELETRONICO'
    ];
    const blocosDeDados = {};

    abasConsultadas.forEach(nomeAba => {
      const sheet = ss.getSheetByName(nomeAba);
      blocosDeDados[nomeAba] = sheet ? sheet.getDataRange().getValues() : [];
    });

    // 1. VENDAS
    const dVendas = blocosDeDados['HISTORICO_VENDA_ANO'];
    for(let i = 1; i < dVendas.length; i++) {
      let rowF = parseInt(normalizarFilialId(dVendas[i][0]), 10);
      if (rowF === targetFilial) {
        let v = dVendas[i][7];
        if (typeof v === 'number') venda = (v * 100).toFixed(2);
        else { 
          let num = parseFloat(String(v).replace('%', '').replace(',', '.')); 
          if (!isNaN(num)) venda = (num < 10 ? num * 100 : num).toFixed(2); 
        }
        break;
      }
    }

    // 2. NPS
    const dNPS = blocosDeDados['HISTORICO_NPS'];
    let maxDateNPS = 0; 
    let lastNPS = "-";
    for(let i = 1; i < dNPS.length; i++) {
      let rowF = parseInt(normalizarFilialId(dNPS[i][1]), 10);
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dNPS[i][0]);
        if (currentDt >= maxDateNPS) {
          maxDateNPS = currentDt;
          let v = parseFloat(String(dNPS[i][6]).replace(',', '.'));
          if(!isNaN(v)) lastNPS = v.toFixed(1);
        }
      }
    }
    nps = lastNPS;

    // 3. BANCO DE HORAS
    const dBH = blocosDeDados['HISTORICO_BH_ACUMULADO'];
    let maxDateBH = 0; 
    let sumBH = 0; 
    let matchFoundBH = false;
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(normalizarFilialId(dBH[i][1]), 10);
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt > maxDateBH) maxDateBH = currentDt;
      }
    }
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(normalizarFilialId(dBH[i][1]), 10);
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt === maxDateBH) {
          let v = parseFloat(String(dBH[i][9]).replace(',', '.'));
          if(!isNaN(v)) { sumBH += v; matchFoundBH = true; }
        }
      }
    }
    if(matchFoundBH) bh = sumBH.toFixed(2);

    // 4. QUADRO (VAGAS)
    const dQuadro = blocosDeDados['HISTORICO_QUADRO'];
    let qContratar = 0; 
    let matchFoundQ = false;
    for(let i = 1; i < dQuadro.length; i++) {
      let rowF = parseInt(normalizarFilialId(dQuadro[i][0]), 10);
      if (rowF === targetFilial) {
        let cargo = String(dQuadro[i][7]).trim();
        if (cargo !== "Intermitente" && cargo !== "Outros - Montagem") {
          let contratarVal = parseFloat(String(dQuadro[i][12]).replace(',', '.'));
          if(!isNaN(contratarVal)) { qContratar += contratarVal; matchFoundQ = true; }
        }
      }
    }
    if(matchFoundQ) quadro = Math.round(qContratar);

    // 5. TAXA DE DESLIGAMENTO (HISTORICO_TXDESL)
    // Coluna A (0) = Ano, Coluna B (1) = filial, Coluna K (10) = Tx de Saída
    const dTx = blocosDeDados['HISTORICO_TXDESL'];
    for(let i = 1; i < dTx.length; i++) {
      let rowF = parseInt(normalizarFilialId(dTx[i][1] !== undefined && dTx[i][1] !== '' ? dTx[i][1] : dTx[i][0]), 10);
      if (rowF === targetFilial) {
        let rawVal = dTx[i][10] !== undefined && dTx[i][10] !== '' ? dTx[i][10] : dTx[i][9];
        let v = parseFloat(String(rawVal).replace('%', '').replace(',', '.'));
        if(!isNaN(v)) {
          let perc = (v < 1 ? v * 100 : v).toFixed(2);
          if (txdes === "-" || parseFloat(perc) > parseFloat(txdes)) txdes = perc;
        }
      }
    }

    // 6. LIXO ELETRÔNICO / COLETAS
    // Busca na DB_DASH primeiro
    ['HISTORICO_LIXO', 'HISTORICO_COLETAS', 'HISTORICO_LIXO_ELETRONICO'].forEach(nomeAbaLixo => {
      const dL = blocosDeDados[nomeAbaLixo];
      if (dL && dL.length > 1) {
        for (let i = 1; i < dL.length; i++) {
          let rowF = parseInt(normalizarFilialId(dL[i][1] !== undefined && dL[i][1] !== '' ? dL[i][1] : dL[i][0]), 10);
          if (rowF === targetFilial) {
            let val = parseFloat(String(dL[i][dL[i].length - 1] || dL[i][7] || dL[i][2]).replace(',', '.'));
            if (!isNaN(val)) lixo = Math.round(val);
          }
        }
      }
    });

    // Fallback de busca de lixo na aba DADOS_INDICADORES do DB_MASTER (Coluna H / Index 7)
    if (lixo === "-") {
      try {
        const ssMaster = SpreadsheetApp.openById(SPREADSHEET_ID);
        const abaInd = ssMaster.getSheetByName('DADOS_INDICADORES');
        if (abaInd) {
          const dadosInd = abaInd.getDataRange().getValues();
          for (let i = 1; i < dadosInd.length; i++) {
            let rowF = parseInt(normalizarFilialId(dadosInd[i][0]), 10);
            if (rowF === targetFilial) {
              let valLixo = parseNumPtBr(dadosInd[i][7]); 
              if (!isNaN(valLixo)) lixo = Math.round(valLixo);
              break;
            }
          }
        }
      } catch (eLixo) {
        Logger.log("Erro no fallback de lixo: " + eLixo.message);
      }
    }

    let payload = {
      sucesso: true,
      venda: venda !== "-" ? String(venda).replace('.', ',') + "%" : "-",
      nps: nps !== "-" ? String(nps).replace('.', ',') : "-",
      bancoHoras: bh !== "-" ? String(bh).replace('.', ',') : "-",
      lixo: lixo !== "-" ? String(lixo) : "-",
      quadro: quadro !== "-" ? String(quadro) : "-",
      txdes: txdes !== "-" ? String(txdes).replace('.', ',') + "%" : "-"
    };
    cache.put(cacheKey, JSON.stringify(payload), 300);
    return payload;
  } catch(e) { return { erro: e.message }; }
}

function carregarNaturezasSeguras(ss) {
  let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
  if (!sheetConfig) return [];
  const valores = sheetConfig.getDataRange().getValues();
  const fixas = ["Reunião regional", "Treinamentos", "Celebrações/Ritão/Reconhecimento", "Recrutamento e seleção - 1º liderança", "Recrutamento e seleção - Processo Externo", "Reunião Conselho E Conselho consultivo", "NPS - Lojas", "Receita de Mercadoria", "GMD - Operações de Loja", "Relatorios/Feedback/Acompanhamento", "Atendimento Social"];
  let list = [];
  for (let i = 1; i < valores.length; i++) {
    let item = valores[i][0] ? String(valores[i][0]).trim() : "";
    if (item && item !== 'Naturezas_Atividades' && !fixas.includes(item)) list.push(item);
  }
  return list;
}
