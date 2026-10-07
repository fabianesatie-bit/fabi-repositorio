/**
 * ===================================================================
 * SISTEMA GP MAGALU - DASHBOARD DE PERFORMANCE & FEEDBACK
 * Arquivo: Code.gs (Versão com Proteção de Tamanho de Carga)
 * ===================================================================
 */

var DASHBOARD_CONFIG = {
  DB_MASTER_ID: "1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA",
  SHEETS: {
    FEEDBACK: 'REGISTRO_FEEDBACK',
    HISTORICO: 'HISTORICO1',
    FILIAIS: 'Base filiais'
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard de Performance & Feedback | Magalu')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetH = ss.getSheetByName(DASHBOARD_CONFIG.SHEETS.HISTORICO);
    var sheetF = ss.getSheetByName(DASHBOARD_CONFIG.SHEETS.FEEDBACK);

    if (!sheetH) throw new Error("Aba '" + DASHBOARD_CONFIG.SHEETS.HISTORICO + "' não encontrada.");

    var dadosH = sheetH.getDataRange().getValues();
    if (dadosH.length <= 1) return JSON.stringify({});

    // 1. Mapeia Vendas e Metas de REGISTRO_FEEDBACK
    var mapaVendasMetas = {};
    var mapaRegionaisFiliais = {};
    var mapaDiretorias = {};

    if (sheetF) {
      var dadosF = sheetF.getDataRange().getValues();
      var rawHeadersF = dadosF.shift();
      var hMapF = mapearCabecalhos(rawHeadersF);

      var idxF = {
        id: encontrarColuna(hMapF, ['ID', 'CPF', 'MATRICULA']),
        venda: encontrarColuna(hMapF, ['VENDA1P', 'VENDA']),
        meta: encontrarColuna(hMapF, ['META1P', 'META']),
        regional: encontrarColuna(hMapF, ['REGIONAL']),
        diretoria: encontrarColuna(hMapF, ['DIRETORIA']),
        filial: encontrarColuna(hMapF, ['FILIAL'])
      };

      dadosF.forEach(function(r) {
        var id = normalizarID(r[idxF.id]);
        var fil = normalizarFilialId(r[idxF.filial]);
        var reg = String(r[idxF.regional] || 'Sem Regional').trim();
        var dir = idxF.diretoria !== -1 ? String(r[idxF.diretoria] || 'Sem Diretoria').trim() : 'Geral';

        if (id) {
          mapaVendasMetas[id] = {
            venda: Number(r[idxF.venda]) || 0,
            meta: Number(r[idxF.meta]) || 0
          };
        }
        if (fil) {
          mapaRegionaisFiliais[fil] = reg;
          mapaDiretorias[fil] = dir;
        }
      });
    }

    // 2. Carrega telefones dos gerentes da DB_MASTER
    var mapaTelefones = obterMapaTelefonesGerentes();

    var dadosBrutos = {
      global: [],
      diretorias: {},
      regionais: {},
      filiais: {},
      mapa: {},
      mapaRegionalFilial: {}
    };

    var hoje = new Date();
    hoje.setHours(0,0,0,0);

    for (var i = 1; i < dadosH.length; i++) {
      var row = dadosH[i];
      var id = normalizarID(row[0]);
      var nome = String(row[1] || 'Sem Nome').trim();
      var filial = normalizarFilialId(row[2]);
      var dtAdmStr = row[3];
      var f15 = String(row[4] || '').toUpperCase();
      var f30 = String(row[5] || '').toUpperCase();
      var f45 = String(row[6] || '').toUpperCase();
      var statusProc = String(row[7] || '').toUpperCase();
      var diasAtrasoRegistrado = Number(row[8]) || 0;

      // Lê diretamente da aba limpa HISTORICO1 (Ignora quem está desligado)
      if (!id || statusProc === "DESLIGADO" || !filial) continue;

      var dtAdm = parseDataBR(dtAdmStr);
      var diasCasa = dtAdm ? calcularDiasCasa(dtAdm, hoje) : 0;
      if (diasCasa < 15) continue;

      var regional = mapaRegionaisFiliais[filial] || "Sem Regional";
      var diretoria = mapaDiretorias[filial] || "Geral";
      var financeiro = mapaVendasMetas[id] || { venda: 0, meta: 0 };
      var isPendente = (diasCasa >= 15 && f15 !== "OK") || (diasCasa >= 30 && f30 !== "OK") || (diasCasa >= 45 && f45 !== "OK");

      var item = {
        id: id,
        nome: nome,
        filial: filial,
        diasCasa: diasCasa,
        f15: f15,
        f30: f30,
        f45: f45,
        isPendente: isPendente,
        diasAtraso: diasAtrasoRegistrado,
        venda: financeiro.venda,
        meta: financeiro.meta
      };

      if (!dadosBrutos.mapa[diretoria]) dadosBrutos.mapa[diretoria] = new Set();
      dadosBrutos.mapa[diretoria].add(regional);

      if (!dadosBrutos.mapaRegionalFilial[regional]) dadosBrutos.mapaRegionalFilial[regional] = new Set();
      dadosBrutos.mapaRegionalFilial[regional].add(filial);

      dadosBrutos.global.push(item);
      (dadosBrutos.diretorias[diretoria] = dadosBrutos.diretorias[diretoria] || []).push(item);
      (dadosBrutos.regionais[regional] = dadosBrutos.regionais[regional] || []).push(item);
      (dadosBrutos.filiais[filial] = dadosBrutos.filiais[filial] || []).push(item);
    }

    // 3. Processamento de Métricas
    var processarMetricas = function(itens) {
      var somaVendasEmDia = 0, somaMetasEmDia = 0;
      var somaVendasPendentes = 0, somaMetasPendentes = 0;
      var pendentesList = [];
      var totalConcluidos = 0;

      itens.forEach(function(it) {
        if (!it.isPendente) {
          totalConcluidos++;
          somaVendasEmDia += it.venda;
          somaMetasEmDia += it.meta;
        } else {
          somaVendasPendentes += it.venda;
          somaMetasPendentes += it.meta;

          var tags = [];
          if (it.diasCasa >= 15 && it.f15 !== "OK") tags.push('FB 15 Dias');
          if (it.diasCasa >= 30 && it.f30 !== "OK") tags.push('FB 30 Dias');
          if (it.diasCasa >= 45 && it.f45 !== "OK") tags.push('FB 45 Dias');

          pendentesList.push({
            nome: it.nome,
            filial: it.filial,
            diasContratados: it.diasCasa,
            quaisPendentes: tags.join(', '),
            diasAtraso: it.diasAtraso,
            atingimento: it.meta > 0 ? (it.venda / it.meta) * 100 : 0
          });
        }
      });

      var totalVendedores = itens.length;
      var mediaAtingimentoEmDia = somaMetasEmDia > 0 ? (somaVendasEmDia / somaMetasEmDia) * 100 : 0;
      var mediaAtingimentoPendente = somaMetasPendentes > 0 ? (somaVendasPendentes / somaMetasPendentes) * 100 : 0;
      var pctEmDia = totalVendedores > 0 ? (totalConcluidos / totalVendedores) * 100 : 0;

      var faturamentoPotencialPendentes = somaMetasPendentes * (mediaAtingimentoEmDia / 100);
      var faturamentoPerdido = Math.max(0, faturamentoPotencialPendentes - somaVendasPendentes);
      var somaVendasAtual = somaVendasEmDia + somaVendasPendentes;
      var incrementoPct = (somaVendasAtual > 0 && faturamentoPerdido > 0) ? (faturamentoPerdido / somaVendasAtual) * 100 : 0;

      return {
        mediaAtingimentoEmDia: mediaAtingimentoEmDia,
        mediaAtingimentoPendente: mediaAtingimentoPendente,
        somaVendasAtual: somaVendasAtual,
        faturamentoPerdido: faturamentoPerdido,
        incrementoPotencialPercentual: incrementoPct,
        totalVendedores: totalVendedores,
        totalPendentes: pendentesList.length,
        totalConcluidos: totalConcluidos,
        pctEmDia: pctEmDia,
        colaboradoresPendentes: pendentesList.sort(function(a, b) { return b.diasAtraso - a.diasAtraso; })
      };
    };

    var mapToObjectArray = function(m) {
      var obj = {};
      for (var k in m) obj[k] = Array.from(m[k]);
      return obj;
    };

    var finalResult = {
      global: processarMetricas(dadosBrutos.global),
      diretorias: {},
      regionais: {},
      filiais: {},
      mapa: mapToObjectArray(dadosBrutos.mapa),
      mapaRegionalFilial: mapToObjectArray(dadosBrutos.mapaRegionalFilial)
    };

    for (var d in dadosBrutos.diretorias) {
      finalResult.diretorias[d] = processarMetricas(dadosBrutos.diretorias[d]);
    }

    for (var r in dadosBrutos.regionais) {
      var metReg = processarMetricas(dadosBrutos.regionais[r]);
      var dirPai = "Geral";
      for (var dirKey in dadosBrutos.mapa) {
        if (dadosBrutos.mapa[dirKey].has(r)) { dirPai = dirKey; break; }
      }
      metReg.diretoria = dirPai;
      finalResult.regionais[r] = metReg;
    }

    for (var f in dadosBrutos.filiais) {
      var metFil = processarMetricas(dadosBrutos.filiais[f]);
      metFil.telefoneGerente = mapaTelefones[f] || "";
      var regPai = "Sem Regional";
      for (var regKey in dadosBrutos.mapaRegionalFilial) {
        if (dadosBrutos.mapaRegionalFilial[regKey].has(f)) { regPai = regKey; break; }
      }
      metFil.regional = regPai;

      var dirFil = "Geral";
      for (var dKey in dadosBrutos.mapa) {
        if (dadosBrutos.mapa[dKey].has(regPai)) { dirFil = dKey; break; }
      }
      metFil.diretoria = dirFil;
      finalResult.filiais[f] = metFil;
    }

    return JSON.stringify(finalResult);

  } catch (error) {
    return JSON.stringify({ _error: true, message: error.message });
  }
}

function obterMapaTelefonesGerentes() {
  var mapa = {};
  try {
    var ss = SpreadsheetApp.openById(DASHBOARD_CONFIG.DB_MASTER_ID);
    var sheet = ss.getSheetByName("DADOS_LOJAS");
    if (!sheet) return mapa;

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return mapa;

    var headers = data[0].map(function(h) { return String(h || '').toUpperCase().trim(); });
    var idxFilial = headers.indexOf("FILIAL_ID");
    if (idxFilial === -1) idxFilial = headers.indexOf("FILIAL");
    if (idxFilial === -1) idxFilial = 0;

    var idxTel = headers.indexOf("TELEFONE_GERENTE");
    if (idxTel === -1) idxTel = headers.indexOf("TELEFONE");
    if (idxTel === -1) idxTel = 7;

    for (var i = 1; i < data.length; i++) {
      var filRaw = data[i][idxFilial];
      var telRaw = data[i][idxTel];
      if (filRaw !== "" && filRaw !== null && filRaw !== undefined) {
        var fKey = normalizarFilialId(filRaw);
        var tLimp = String(telRaw || "").replace(/\D/g, "");
        if (fKey) mapa[fKey] = tLimp;
      }
    }
  } catch (err) {}
  return mapa;
}

function salvarTelefoneGerente(filialId, telefone) {
  try {
    var idNormalizado = normalizarFilialId(filialId);
    if (!idNormalizado) return { sucesso: false, mensagem: "ID de Filial inválido." };

    var ss = SpreadsheetApp.openById(DASHBOARD_CONFIG.DB_MASTER_ID);
    var sheet = ss.getSheetByName("DADOS_LOJAS");
    if (!sheet) {
      sheet = ss.insertSheet("DADOS_LOJAS");
      sheet.appendRow(["FILIAL_ID", "NOME_LOJA", "REGIONAL", "DIRETORIA", "GERENTE", "EMAIL", "UF", "TELEFONE_GERENTE"]);
    }

    var data = sheet.getDataRange().getValues();
    var telefoneLimpo = String(telefone || "").replace(/\D/g, "");

    var idxFilial = 0;
    var idxTel = 7;
    if (data.length > 0) {
      var headers = data[0].map(function(h) { return String(h || '').toUpperCase().trim(); });
      var fIdx = headers.indexOf("FILIAL_ID");
      if (fIdx !== -1) idxFilial = fIdx;
      var tIdx = headers.indexOf("TELEFONE_GERENTE");
      if (tIdx !== -1) idxTel = tIdx;
    }

    var linhaEncontrada = -1;
    for (var i = 1; i < data.length; i++) {
      if (normalizarFilialId(data[i][idxFilial]) === idNormalizado) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada !== -1) {
      sheet.getRange(linhaEncontrada, idxTel + 1).setValue(telefoneLimpo);
    } else {
      var novaLinha = new Array(Math.max(idxTel + 1, 8)).fill("");
      novaLinha[idxFilial] = idNormalizado;
      novaLinha[idxTel] = telefoneLimpo;
      sheet.appendRow(novaLinha);
    }

    return { sucesso: true, mensagem: "Telefone salvo com sucesso!" };
  } catch (error) {
    return { sucesso: false, mensagem: error.toString() };
  }
}

function normalizarTexto(t) {
  return t ? t.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim() : "";
}

function normalizarID(id) {
  if (id === undefined || id === null) return "";
  var str = String(id).trim().replace(/^0+/, '');
  return str === "" ? "0" : str;
}

function normalizarFilialId(filialId) {
  if (filialId === null || filialId === undefined) return "";
  return String(filialId).trim().replace(/^0+/, "");
}

function parseDataBR(s) {
  if (!s) return null;
  if (s instanceof Date) return new Date(s.getFullYear(), s.getMonth(), s.getDate());
  var str = s.toString().trim();
  var p = str.split("/");
  if (p.length === 3) return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
  var p2 = str.split("-");
  if (p2.length === 3) return new Date(parseInt(p2[0], 10), parseInt(p2[1], 10) - 1, parseInt(p2[2], 10));
  return null;
}

function calcularDiasCasa(dataAdmissao, hoje) {
  if (!dataAdmissao) return -1;
  var dAdm = new Date(dataAdmissao.getFullYear(), dataAdmissao.getMonth(), dataAdmissao.getDate());
  var dHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  return Math.round((dHoje.getTime() - dAdm.getTime()) / (1000 * 60 * 60 * 24));
}

function mapearCabecalhos(headerRow) {
  var map = {};
  headerRow.forEach(function(col, idx) { map[normalizarTexto(col).replace(/[\s_]+/g, '')] = idx; });
  return map;
}

function encontrarColuna(headerMap, candidatos) {
  for (var i = 0; i < candidatos.length; i++) {
    var clean = normalizarTexto(candidatos[i]).replace(/[\s_]+/g, '');
    for (var key in headerMap) {
      if (key.includes(clean)) return headerMap[key];
    }
  }
  return -1;
}
