// =============================================================================
// CARREGAMENTO DE DADOS E RELACIONAMENTOS (PORTAL HOME & DASHBOARDS)
// =============================================================================

function obterDadosIniciais() {
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
    let gastoTotalAcumulado = 0;
    let setVisitasFisicas = new Set();

    const dHoje = new Date();
    const mesAtual = dHoje.getMonth();
    const anoAtual = dHoje.getFullYear();
    const meusLancamentos = []; 
    const sheetLancamentos = ss.getSheetByName('DADOS_LANCAMENTOS'); 

    if (sheetLancamentos && sheetLancamentos.getLastRow() > 1) {
      const dadosLancamentos = sheetLancamentos.getDataRange().getValues();
      for (let i = 1; i < dadosLancamentos.length; i++) {
        const emailAutor = dadosLancamentos[i][2] ? String(dadosLancamentos[i][2]).toLowerCase().trim() : "";
        if (emailAutor) {
           if (!placarMoedasTotal[emailAutor]) placarMoedasTotal[emailAutor] = 0;
           if (!placarMoedasMes[emailAutor]) placarMoedasMes[emailAutor] = 0;

           let valorMoedaRaw = String(dadosLancamentos[i][9] || "0").replace(',', '.').trim();
           let moedasParse = parseFloat(valorMoedaRaw);
           if (!isNaN(moedasParse)) {
               placarMoedasTotal[emailAutor] += moedasParse;

               let dtAcaoRaw = dadosLancamentos[i][12] || dadosLancamentos[i][1];
               if (dtAcaoRaw) {
                 let dtObj = new Date(dtAcaoRaw);
                 if (!isNaN(dtObj.getTime()) && dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual) {
                   placarMoedasMes[emailAutor] += moedasParse;
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
                  if (dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual) {
                      gastoTotalAcumulado += gastoSomaReal;
                  }
              }
              let rot = String(dadosLancamentos[i][17] || "").trim();
              if (rot.includes("Visita in loco")) {
                  let dtVFormat = formatarDataSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
                  let dst = String(dadosLancamentos[i][3]).trim();
                  setVisitasFisicas.add(emailAutor + "|" + dtVFormat + "|" + dst);
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
    let historicoLimitado = meusLancamentos.slice(0, 500);

    let premiosDicionario = getDicionarioPremios();
    let userMoedasTotal = placarMoedasTotal[emailLogado] || 0;
    let userMoedasMes = placarMoedasMes[emailLogado] || 0;

    let badgesConquistadas = premiosDicionario.filter(p => userMoedasTotal >= p.meta && p.meta > 0);

    let faseMontanha = 1;
    if (userMoedasTotal >= 120) faseMontanha = 4;
    else if (userMoedasTotal >= 80) faseMontanha = 3;
    else if (userMoedasTotal >= 40) faseMontanha = 2;

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
        let badgesU = premiosDicionario.filter(p => totalU >= p.meta && p.meta > 0).map(p => p.icone);

        let faseU = 1;
        if (totalU >= 120) faseU = 4;
        else if (totalU >= 80) faseU = 3;
        else if (totalU >= 40) faseU = 2;

        return { 
          email: email, 
          nome: nomeExibicao, 
          foto: fotoExibicao, 
          moedas: totalU,
          moedasMes: placarMoedasMes[email] || 0,
          fase: faseU,
          badges: badgesU
        };
    });
    ranking = ranking.filter(user => user.moedas > 0 && !isNaN(user.moedas));
    ranking.sort((a, b) => b.moedas - a.moedas);

    const lojasFiltradasMap = new Map();
    const sheetLojas = ss.getSheetByName('DADOS_LOJAS');
    if (sheetLojas && sheetLojas.getLastRow() > 1) {
      const dadosLojas = sheetLojas.getDataRange().getValues();
      for (let i = 1; i < dadosLojas.length; i++) {
        const regionalLoja = dadosLojas[i][2] ? String(dadosLojas[i][2]).trim() : "";
        if (usuario.isSuperAdmin || usuario.regionais.includes(regionalLoja)) {
          let idLoja = parseInt(dadosLojas[i][0], 10);
          if (!isNaN(idLoja)) {
              if (idLoja > 3000) idLoja = idLoja - 3000;
              if (!lojasFiltradasMap.has(idLoja)) {
                  lojasFiltradasMap.set(idLoja, { id: idLoja, nome: dadosLojas[i][1], regional: regionalLoja });
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
      qtdVisitas: setVisitasFisicas.size,
      moedas: userMoedasTotal,
      moedasMes: userMoedasMes,
      faseMontanha: faseMontanha,
      badges: badgesConquistadas,
      gastos: gastoTotalAcumulado, historico: historicoLimitado, 
      avisos: avisos, ranking: ranking.slice(0, 5), naturezas: naturezas, diretorias: diretoriasAr
    };
  } catch (e) { return { erro: e.message }; }
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

function buscarIndicadoresLoja(filialId) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const verificacao = validarAcessoFilial(emailLogado, filialId);
    if (!verificacao.autorizado) {
      return { erro: "ACESSO NEGADO: Você não tem permissão para ler dados desta filial." };
    }
    let targetFilial = parseInt(filialId, 10);
    if (targetFilial > 3000) targetFilial -= 3000;
    const cache = CacheService.getScriptCache();
    const cacheKey = 'IND_LOJA_' + targetFilial;
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
    const ss = SpreadsheetApp.openById(SPREADSHEET_DASH);
    let nps = "-", venda = "-", bh = "-", quadro = "-", txdes = "-";
    const abasConsultadas = ['HISTORICO_VENDA_ANO', 'HISTORICO_NPS', 'HISTORICO_BH_ACUMULADO', 'HISTORICO_QUADRO', 'HISTORICO_TXDESL'];
    const blocosDeDados = {};

    abasConsultadas.forEach(nomeAba => {
      const sheet = ss.getSheetByName(nomeAba);
      blocosDeDados[nomeAba] = sheet ? sheet.getDataRange().getValues() : [];
    });

    const dVendas = blocosDeDados['HISTORICO_VENDA_ANO'];
    for(let i = 1; i < dVendas.length; i++) {
      let rowF = parseInt(dVendas[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
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

    const dNPS = blocosDeDados['HISTORICO_NPS'];
    let maxDateNPS = 0; 
    let lastNPS = "-";
    for(let i = 1; i < dNPS.length; i++) {
      let rowF = parseInt(dNPS[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
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

    const dBH = blocosDeDados['HISTORICO_BH_ACUMULADO'];
    let maxDateBH = 0; 
    let sumBH = 0; 
    let matchFoundBH = false;
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(dBH[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt > maxDateBH) maxDateBH = currentDt;
      }
    }
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(dBH[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt === maxDateBH) {
          let v = parseFloat(String(dBH[i][9]).replace(',', '.'));
          if(!isNaN(v)) { sumBH += v; matchFoundBH = true; }
        }
      }
    }
    if(matchFoundBH) bh = sumBH.toFixed(2);

    const dQuadro = blocosDeDados['HISTORICO_QUADRO'];
    let qContratar = 0; 
    let matchFoundQ = false;
    for(let i = 1; i < dQuadro.length; i++) {
      let rowF = parseInt(dQuadro[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let cargo = String(dQuadro[i][7]).trim();
        if (cargo !== "Intermitente" && cargo !== "Outros - Montagem") {
          let contratarVal = parseFloat(String(dQuadro[i][12]).replace(',', '.'));
          if(!isNaN(contratarVal)) { qContratar += contratarVal; matchFoundQ = true; }
        }
      }
    }
    if(matchFoundQ) quadro = Math.round(qContratar);

    const dTx = blocosDeDados['HISTORICO_TXDESL'];
    for(let i = 1; i < dTx.length; i++) {
      let rowF = parseInt(dTx[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let v = parseFloat(String(dTx[i][9]).replace('%', '').replace(',', '.'));
        if(!isNaN(v)) {
          let perc = (v * 100).toFixed(2);
          if (txdes === "-" || parseFloat(perc) > parseFloat(txdes)) txdes = perc;
        }
      }
    }
    let payload = {
      sucesso: true,
      venda: venda !== "-" ? String(venda).replace('.', ',') + "%" : "-",
      nps: nps !== "-" ? String(nps).replace('.', ',') : "-",
      bancoHoras: bh !== "-" ? String(bh).replace('.', ',') : "-",
      quadro: quadro !== "-" ? String(quadro) : "-",
      txdes: txdes !== "-" ? String(txdes).replace('.', ',') + "%" : "-"
    };
    cache.put(cacheKey, JSON.stringify(payload), 300);
    return payload;
  } catch(e) { return { erro: e.message }; }
}
