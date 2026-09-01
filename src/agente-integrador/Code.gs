/**
 * Inicializa a aplicacao Web App renderizando o arquivo Index.html.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard de Performance & Feedback | Magalu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Middleware obrigatório para normalizar o ID da filial (Remove zeros à esquerda e caracteres não numéricos).
 * @param {string|number} filialId
 * @return {string}
 */
function normalizarFilialId(filialId) {
  if (filialId === null || filialId === undefined) return "";
  const limpo = String(filialId).trim().replace(/^0+/, "");
  return limpo;
}

/**
 * ID da Planilha Mestre de Dados de Lojas
 */
const DB_MASTER_ID = "1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA";

/**
 * Leitura resiliente dos telefones gravados na aba DADOS_LOJAS da DB_MASTER.
 * @return {Object} Mapa de filialIdNormalizado -> Telefone
 */
function obterMapaTelefonesGerentes() {
  const mapa = {};
  try {
    const ss = SpreadsheetApp.openById(DB_MASTER_ID);
    const sheet = ss.getSheetByName("DADOS_LOJAS");
    if (!sheet) return mapa;

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return mapa;

    const headers = data[0].map(h => String(h || '').toUpperCase().trim());
    
    // Tenta localizar as colunas por nome ou assume padrões (Col A = Filial, Col H = Telefone)
    let idxFilial = headers.indexOf("FILIAL_ID");
    if (idxFilial === -1) idxFilial = headers.indexOf("FILIAL");
    if (idxFilial === -1) idxFilial = 0; // Padrão Coluna A

    let idxTel = headers.indexOf("TELEFONE_GERENTE");
    if (idxTel === -1) idxTel = headers.indexOf("TELEFONE");
    if (idxTel === -1) idxTel = 7; // Padrão Coluna H (índice 7)

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const filialRaw = row[idxFilial];
      const telRaw = row[idxTel];

      if (filialRaw !== "" && filialRaw !== null && filialRaw !== undefined) {
        const filialKey = normalizarFilialId(filialRaw);
        const telLimpo = String(telRaw || "").replace(/\D/g, "");
        if (filialKey) {
          mapa[filialKey] = telLimpo;
        }
      }
    }
  } catch (err) {
    Logger.log("Aviso/Erro ao ler mapa de telefones: " + err.toString());
  }
  return mapa;
}

/**
 * Salva assincronamente o telefone do gerente na aba DADOS_LOJAS da DB_MASTER.
 * Resiliente contra linhas em branco e colunas inexistentes.
 */
function salvarTelefoneGerente(filialId, telefone) {
  try {
    const idNormalizado = normalizarFilialId(filialId);
    if (!idNormalizado) {
      return { sucesso: false, mensagem: "ID de Filial inválido ou vazio." };
    }

    const ss = SpreadsheetApp.openById(DB_MASTER_ID);
    let sheet = ss.getSheetByName("DADOS_LOJAS");
    
    if (!sheet) {
      sheet = ss.insertSheet("DADOS_LOJAS");
      sheet.appendRow(["FILIAL_ID", "NOME_LOJA", "REGIONAL", "DIRETORIA", "GERENTE", "EMAIL", "UF", "TELEFONE_GERENTE"]);
    }

    const data = sheet.getDataRange().getValues();
    const telefoneLimpo = String(telefone || "").replace(/\D/g, "");

    let idxFilial = 0;
    let idxTel = 7;

    if (data.length > 0) {
      const headers = data[0].map(h => String(h || '').toUpperCase().trim());
      const fIdx = headers.indexOf("FILIAL_ID");
      if (fIdx !== -1) idxFilial = fIdx;
      const tIdx = headers.indexOf("TELEFONE_GERENTE");
      if (tIdx !== -1) idxTel = tIdx;
    }

    let linhaEncontrada = -1;
    for (let i = 1; i < data.length; i++) {
      const cellFilial = data[i][idxFilial];
      if (cellFilial !== "" && cellFilial !== null && cellFilial !== undefined) {
        if (normalizarFilialId(cellFilial) === idNormalizado) {
          linhaEncontrada = i + 1; // 1-indexed no Sheets
          break;
        }
      }
    }

    if (linhaEncontrada !== -1) {
      sheet.getRange(linhaEncontrada, idxTel + 1).setValue(telefoneLimpo);
    } else {
      // Monta linha preservando posições se adicionar nova
      const novaLinha = new Array(Math.max(idxTel + 1, 8)).fill("");
      novaLinha[idxFilial] = idNormalizado;
      novaLinha[idxTel] = telefoneLimpo;
      sheet.appendRow(novaLinha);
    }

    return { sucesso: true, mensagem: "Telefone salvo com sucesso!" };
  } catch (error) {
    Logger.log("Erro em salvarTelefoneGerente: " + error.toString());
    return { sucesso: false, mensagem: error.toString() };
  }
}

/**
 * VERSÃO DEFINITIVA COM SUPORTE COMPLETO A FILTROS E AGRUPAMENTO
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName('REGISTRO_FEEDBACK');
    const baseFiliaisSheet = ss.getSheetByName('Base filiais');
    const desligamentoSheet = ss.getSheetByName('Atualizacao_desligamento');
    
    if (!feedbackSheet) {
      throw new Error('Aba "REGISTRO_FEEDBACK" não foi encontrada.');
    }
    
    const allData = feedbackSheet.getDataRange().getValues();
    if (allData.length <= 1) return JSON.stringify({}); 

    const rawHeaders = allData.shift();
    
    const normalize = (str) => String(str || '')
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();

    const findColumnIndex = (headersList, candidates) => {
      for (let cand of candidates) {
        const normCand = cand.toUpperCase().replace(/[\s_]+/g, '');
        for (let i = 0; i < headersList.length; i++) {
          const normHeader = String(headersList[i] || '').toUpperCase().replace(/[\s_]+/g, '');
          if (normHeader.includes(normCand)) return i;
        }
      }
      return -1;
    };

    const idx = {
      id: findColumnIndex(rawHeaders, ['ID', 'CPF', 'MATRICULA']),
      nome: findColumnIndex(rawHeaders, ['NOME']),
      diretoria: findColumnIndex(rawHeaders, ['DIRETORIA', 'DIRETOR']),
      regional: findColumnIndex(rawHeaders, ['REGIONAL']),
      familiaCargo: findColumnIndex(rawHeaders, ['FAMILIA_CARGO', 'FAMILIA', 'CARGO']),
      diasContratados: findColumnIndex(rawHeaders, ['DIAS_CONTRATADOS', 'DIAS']),
      f15: findColumnIndex(rawHeaders, ['FEEDBACK_15_DIAS', 'FEEDBACK_15', '15']),
      f30: findColumnIndex(rawHeaders, ['FEEDBACK_30_DIAS', 'FEEDBACK_30', '30']),
      f45: findColumnIndex(rawHeaders, ['FEEDBACK_45_DIAS', 'FEEDBACK_45', '45']),
      venda1P: findColumnIndex(rawHeaders, ['VENDA 1P', 'VENDA 1', 'VENDA']),
      meta1P: findColumnIndex(rawHeaders, ['META 1P', 'META 1', 'META']),
      filial: findColumnIndex(rawHeaders, ['FILIAL', 'FILIA']),
      desligamentoCol: findColumnIndex(rawHeaders, ['DESLIGAMENTO', 'STATUS', 'SITUACAO'])
    };

    // 1. FILTRAGEM DE DESLIGADOS
    const idsDesligados = new Set();
    const nomesDesligados = new Set();

    if (desligamentoSheet) {
      const desData = desligamentoSheet.getDataRange().getValues();
      if (desData.length > 1) {
        const desHeaders = desData.shift();
        const idxDesId = findColumnIndex(desHeaders, ['ID', 'CPF', 'MATRICULA']);
        const idxDesNome = findColumnIndex(desHeaders, ['NOME', 'COLABORADOR']);
        const idxDesStatus = findColumnIndex(desHeaders, ['DESLIGAMENTO', 'STATUS', 'SITUACAO']);

        desData.forEach(r => {
          const statusVal = idxDesStatus !== -1 ? normalize(r[idxDesStatus]) : 'DESLIGADO';
          const idVal = idxDesId !== -1 ? normalize(r[idxDesId]) : '';
          const nomeVal = idxDesNome !== -1 ? normalize(r[idxDesNome]) : '';

          if (statusVal.includes('DESLIGADO')) {
            if (idVal) idsDesligados.add(idVal);
            if (nomeVal) nomesDesligados.add(nomeVal);
          }
        });
      }
    }

    const mapaTelefonesGerentes = obterMapaTelefonesGerentes();

    const dadosBrutos = {
      global: [],
      diretorias: {},
      regionais: {},
      filiais: {},
      mapa: {},
      mapaRegionalFilial: {}
    };

    // 2. SEPARAÇÃO DE LINHAS COM RELACIONAMENTO
    allData.forEach(row => {
      const idColab = idx.id !== -1 ? normalize(row[idx.id]) : '';
      const nomeColab = idx.nome !== -1 ? normalize(row[idx.nome]) : '';
      const statusColab = idx.desligamentoCol !== -1 ? normalize(row[idx.desligamentoCol]) : '';

      const isDesligado = statusColab.includes('DESLIGADO') || 
                          (idColab && idsDesligados.has(idColab)) || 
                          (nomeColab && nomesDesligados.has(nomeColab));

      if (isDesligado) return;

      const dias = Number(row[idx.diasContratados]) || 0;
      const cargo = String(row[idx.familiaCargo] || '').trim();
      const isVendedor = cargo.toUpperCase().includes('VENDEDOR');
      
      if (isVendedor && dias >= 15) {
        let regional = String(row[idx.regional] || 'Sem Regional').trim();
        const diretoria = idx.diretoria !== -1 ? String(row[idx.diretoria] || 'Sem Diretoria').trim() : 'Geral';
        const filial = normalizarFilialId(row[idx.filial]);
        
        if (!filial) return;

        if (!dadosBrutos.mapa[diretoria]) dadosBrutos.mapa[diretoria] = new Set();
        dadosBrutos.mapa[diretoria].add(regional);

        if (!dadosBrutos.mapaRegionalFilial[regional]) dadosBrutos.mapaRegionalFilial[regional] = new Set();
        dadosBrutos.mapaRegionalFilial[regional].add(filial);

        dadosBrutos.global.push(row);
        
        if (!dadosBrutos.diretorias[diretoria]) dadosBrutos.diretorias[diretoria] = [];
        dadosBrutos.diretorias[diretoria].push(row);
        
        if (!dadosBrutos.regionais[regional]) dadosBrutos.regionais[regional] = [];
        dadosBrutos.regionais[regional].push(row);

        if (!dadosBrutos.filiais[filial]) dadosBrutos.filiais[filial] = [];
        dadosBrutos.filiais[filial].push(row);
      }
    });

    // 3. CÁLCULO E MONTAGEM DA HIERARQUIA
    const processarMetricas = (rows) => {
      let somaVendasEmDia = 0, somaMetasEmDia = 0;
      let somaVendasPendentes = 0, somaMetasPendentes = 0;
      let colaboradoresPendentes = [];
      let totalConcluidos = 0;

      rows.forEach(row => {
        const dias = Number(row[idx.diasContratados]) || 0;
        const f15 = Number(row[idx.f15]) === 1;
        const f30 = Number(row[idx.f30]) === 1;
        const f45 = Number(row[idx.f45]) === 1;
        const isPendente = (dias >= 15 && !f15) || (dias >= 30 && !f30) || (dias >= 45 && !f45);
        const venda = Number(row[idx.venda1P]) || 0;
        const meta = Number(row[idx.meta1P]) || 0;
        
        const filialKey = normalizarFilialId(row[idx.filial]);

        if (!isPendente) {
          totalConcluidos++;
          somaVendasEmDia += venda;
          somaMetasEmDia += meta;
        } else {
          somaVendasPendentes += venda;
          somaMetasPendentes += meta;
          
          let pendenciasStr = [];
          let diasAtraso = 0;
          if (dias >= 15 && !f15) { pendenciasStr.push('FB 15 Dias'); diasAtraso = Math.max(diasAtraso, dias - 15); }
          if (dias >= 30 && !f30) { pendenciasStr.push('FB 30 Dias'); diasAtraso = Math.max(diasAtraso, dias - 30); }
          if (dias >= 45 && !f45) { pendenciasStr.push('FB 45 Dias'); diasAtraso = Math.max(diasAtraso, dias - 45); }
          
          colaboradoresPendentes.push({
            nome: String(row[idx.nome] || 'Colaborador Sem Nome').trim(),
            filial: filialKey,
            diasContratados: dias,
            quaisPendentes: pendenciasStr.join(', '),
            diasAtraso: diasAtraso,
            atingimento: meta > 0 ? (venda / meta) * 100 : 0
          });
        }
      });

      const totalVendedores = rows.length;
      const mediaAtingimentoEmDia = somaMetasEmDia > 0 ? (somaVendasEmDia / somaMetasEmDia) * 100 : 0;
      const mediaAtingimentoPendente = somaMetasPendentes > 0 ? (somaVendasPendentes / somaMetasPendentes) * 100 : 0;
      const pctEmDia = totalVendedores > 0 ? (totalConcluidos / totalVendedores) * 100 : 0;
      
      const faturamentoPotencialPendentes = somaMetasPendentes * (mediaAtingimentoEmDia / 100);
      let faturamentoPerdido = faturamentoPotencialPendentes - somaVendasPendentes;
      if (faturamentoPerdido < 0) faturamentoPerdido = 0; 

      const somaVendasAtual = somaVendasEmDia + somaVendasPendentes;
      let incrementoPotencialPercentual = 0;
      if (somaVendasAtual > 0 && faturamentoPerdido > 0) {
          incrementoPotencialPercentual = (faturamentoPerdido / somaVendasAtual) * 100;
      }

      return {
        mediaAtingimentoEmDia: mediaAtingimentoEmDia,
        mediaAtingimentoPendente: mediaAtingimentoPendente,
        somaVendasAtual: somaVendasAtual,
        faturamentoPerdido: faturamentoPerdido,
        incrementoPotencialPercentual: incrementoPotencialPercentual,
        totalVendedores: totalVendedores,
        totalPendentes: colaboradoresPendentes.length,
        totalConcluidos: totalConcluidos,
        pctEmDia: pctEmDia,
        colaboradoresPendentes: colaboradoresPendentes.sort((a,b) => b.diasAtraso - a.diasAtraso)
      };
    };

    const mapToObjectArray = (map) => {
      const obj = {};
      for (const key in map) obj[key] = Array.from(map[key]);
      return obj;
    };

    const finalResult = {
      global: processarMetricas(dadosBrutos.global),
      diretorias: {},
      regionais: {},
      filiais: {},
      mapa: mapToObjectArray(dadosBrutos.mapa),
      mapaRegionalFilial: mapToObjectArray(dadosBrutos.mapaRegionalFilial)
    };

    for (const dir in dadosBrutos.diretorias) {
      finalResult.diretorias[dir] = processarMetricas(dadosBrutos.diretorias[dir]);
    }
    
    for (const reg in dadosBrutos.regionais) {
      const metricasReg = processarMetricas(dadosBrutos.regionais[reg]);
      let dirPai = "Geral";
      for (let d in dadosBrutos.mapa) {
        if (dadosBrutos.mapa[d].has(reg)) { dirPai = d; break; }
      }
      metricasReg.diretoria = dirPai;
      finalResult.regionais[reg] = metricasReg;
    }

    for (const fil in dadosBrutos.filiais) {
      const metricasFil = processarMetricas(dadosBrutos.filiais[fil]);
      metricasFil.telefoneGerente = mapaTelefonesGerentes[fil] || "";
      
      let regPai = "Sem Regional";
      for (let r in dadosBrutos.mapaRegionalFilial) {
        if (dadosBrutos.mapaRegionalFilial[r].has(fil)) { regPai = r; break; }
      }
      metricasFil.regional = regPai;

      let dirPai = "Geral";
      for (let d in dadosBrutos.mapa) {
        if (dadosBrutos.mapa[d].has(regPai)) { dirPai = d; break; }
      }
      metricasFil.diretoria = dirPai;

      finalResult.filiais[fil] = metricasFil;
    }

    return JSON.stringify(finalResult);

  } catch (error) {
    return JSON.stringify({ _error: true, message: error.message });
  }
}
