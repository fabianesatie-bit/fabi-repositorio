/**
 * ===================================================================
 * SISTEMA GP MAGALU - GESTÃO INTEGRADA (ADMISSÃO + FEEDBACKS + DASHBOARD)
 * Versão 8.0 - Agente Integrador Unificado com Tríplice Regra de Status
 * ===================================================================
 */

const CONFIG = {
  CACHE_KEY: 'AGENTE_INTEGRADOR_DASHBOARD_CACHE_V4',
  CACHE_TTL_SECONDS: 1800,
  SHEETS: {
    FEEDBACK: 'REGISTRO_FEEDBACK',
    ATIVOS: 'DADOS_ATIVOS',
    FILIAIS: 'Base filiais',
    HISTORICO_NOVO: 'HISTORICO1',         // Aba de 9 colunas limpa
    HISTORICO_LEGADO: 'Historico_Legado', // Aba antiga
    ADMISSAO: 'ADMISSAO'
  },
  CARGOS_PILOTO: [
    "VENDEDOR", 
    "ASSISTENTE VENDAS SR", 
    "ASSISTENTE VENDAS JR", 
    "AUX LIMPEZA", 
    "LIDER COMERCIAL", 
    "LIDER DE ESTOQUE"
  ]
};

// ===================================================================
// MENU SUPERIOR NA PLANILHA
// ===================================================================

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('GP Magalu')
      .addItem('🧹 Executar Limpeza e Recalcular Atrasos', 'executarLimpezaInicialHistorico')
      .addSeparator()
      .addItem('🔄 Sincronizar Feedbacks (REGISTRO_FEEDBACK)', 'sincronizarFeedbacksUniversidade')
      .addItem('👥 Processar Desligamentos (DADOS_ATIVOS)', 'processarDesligamentos')
      .addItem('⚡ Limpar Cache do Dashboard', 'limparCacheDashboard')
      .addToUi();
  } catch (e) {
    console.log("Aviso: Menu não carregado.");
  }
}

// ===================================================================
// LIMPEZA INICIAL E RECÁLCULO DOS DIAS EM ATRASO
// ===================================================================

function executarLimpezaInicialHistorico() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAtivos = ss.getSheetByName(CONFIG.SHEETS.ATIVOS);
  var sheetFeedback = ss.getSheetByName(CONFIG.SHEETS.FEEDBACK);
  var sheetLegado = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_LEGADO) || ss.getSheetByName("Historico");
  var sheetNovo = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);

  if (!sheetAtivos || !sheetLegado || !sheetNovo) {
    var msgErro = "Erro: Verifique se as abas 'DADOS_ATIVOS', 'Historico_Legado' e 'HISTORICO1' existem.";
    console.log(msgErro);
    try { SpreadsheetApp.getUi().alert("Erro", msgErro, SpreadsheetApp.getUi().ButtonSet.OK); } catch(e){}
    return;
  }

  // Garante os 9 cabeçalhos na Linha 1
  sheetNovo.getRange(1, 1, 1, 9).setValues([[
    "ID", "NOME", "FILIAL", "DATA_ADMISSAO", "FEEDBACK_15", "FEEDBACK_30", "FEEDBACK_45", "STATUS_PROCESSO", "DIAS_ATRASO"
  ]]);

  var dadosAtivos = sheetAtivos.getDataRange().getValues();
  var headerMap = mapearCabecalhos(dadosAtivos[0]);
  var colId = encontrarColuna(headerMap, ["id", "matricula", "codigo"]);
  var colSituacao = encontrarColuna(headerMap, ["situacao", "status", "situacao_colaborador"]);

  if (colId === -1) colId = 1;       // Fallback: Coluna B
  if (colSituacao === -1) colSituacao = 7; // Fallback: Coluna H

  var ativosSet = new Set();
  var pausadosSet = new Set();

  for (var i = 1; i < dadosAtivos.length; i++) {
    var idAtivo = normalizarID(dadosAtivos[i][colId]);
    var sitAtivo = normalizarTexto(dadosAtivos[i][colSituacao]);
    if (!idAtivo) continue;

    if (sitAtivo.includes("EM ATIVIDADE NORMAL")) {
      ativosSet.add(idAtivo);
    } else if (!sitAtivo.includes("DEMITIDO")) {
      // Afastados, Férias, Licenças, Aposentadoria, etc.
      pausadosSet.add(idAtivo);
    }
  }

  var mapaFeedbacks = {};
  if (sheetFeedback) {
    var dadosFeed = sheetFeedback.getDataRange().getValues();
    for (var f = 1; f < dadosFeed.length; f++) {
      var idFeed = normalizarID(dadosFeed[f][0]);
      if (!idFeed) continue;
      mapaFeedbacks[idFeed] = {
        f15: String(dadosFeed[f][8]) === "1" || normalizarTexto(dadosFeed[f][8]) === "OK",
        f30: String(dadosFeed[f][9]) === "1" || normalizarTexto(dadosFeed[f][9]) === "OK",
        f45: String(dadosFeed[f][10]) === "1" || normalizarTexto(dadosFeed[f][10]) === "OK"
      };
    }
  }

  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  var dadosLegado = sheetLegado.getDataRange().getValues();
  var novasLinhas = [];
  var idsProcessados = new Set();

  var totalDesligadosRemovidos = 0;
  var totalAtivosMantidos = 0;
  var totalPausadosMantidos = 0;

  for (var r = 1; r < dadosLegado.length; r++) {
    var linha = dadosLegado[r];
    var idOriginal = String(linha[3] || linha[0] || '').trim();
    var idLimpo = normalizarID(idOriginal);
    if (!idLimpo || idsProcessados.has(idLimpo)) continue;

    var isAtivo = ativosSet.has(idLimpo);
    var isPausado = pausadosSet.has(idLimpo);

    if (!isAtivo && !isPausado) {
      totalDesligadosRemovidos++;
      continue;
    }

    var nome = String(linha[5] || 'Sem Nome').trim();
    var filial = normalizarFilialId(linha[0]);
    var dataAdmStr = formatarDataBR(linha[6]);
    var dtAdm = parseDataBR(linha[6]);

    var fb = mapaFeedbacks[idLimpo] || { f15: false, f30: false, f45: false };
    var f15Status = (fb.f15 || String(linha[12]).trim().toUpperCase() === "OK") ? "OK" : "PENDENTE";
    var f30Status = (fb.f30 || String(linha[13]).trim().toUpperCase() === "OK") ? "OK" : "PENDENTE";
    var f45Status = (fb.f45 || String(linha[14]).trim().toUpperCase() === "OK") ? "OK" : "PENDENTE";

    var statusProc = "EM ACOMPANHAMENTO";
    if (isPausado) {
      statusProc = "PAUSADO";
      totalPausadosMantidos++;
    } else if (f15Status === "OK" && f30Status === "OK" && f45Status === "OK") {
      statusProc = "CONCLUIDO";
      totalAtivosMantidos++;
    } else {
      totalAtivosMantidos++;
    }

    var diasAtraso = 0;
    if (statusProc === "EM ACOMPANHAMENTO" && dtAdm) {
      var diasCasa = calcularDiasCasa(dtAdm, hoje);
      if (diasCasa >= 15 && f15Status !== "OK") {
        diasAtraso = diasCasa - 15;
      } else if (diasCasa >= 30 && f30Status !== "OK") {
        diasAtraso = diasCasa - 30;
      } else if (diasCasa >= 45 && f45Status !== "OK") {
        diasAtraso = diasCasa - 45;
      }
    }

    novasLinhas.push([
      idOriginal,
      nome,
      filial,
      dataAdmStr,
      f15Status,
      f30Status,
      f45Status,
      statusProc,
      diasAtraso
    ]);

    idsProcessados.add(idLimpo);
  }

  if (sheetNovo.getLastRow() > 1) {
    sheetNovo.getRange(2, 1, sheetNovo.getLastRow() - 1, 9).clearContent();
  }

  if (novasLinhas.length > 0) {
    sheetNovo.getRange(2, 1, novasLinhas.length, 9).setValues(novasLinhas);
  }

  SpreadsheetApp.flush();

  var relatorio = "✅ Higienização e Recálculo Concluídos!\n\n" +
                  "❌ Desligados removidos: " + totalDesligadosRemovidos + "\n" +
                  "👥 Ativos em acompanhamento: " + totalAtivosMantidos + "\n" +
                  "⏸️ Pausados (Férias/Afastados): " + totalPausadosMantidos;

  console.log(relatorio);
  try {
    SpreadsheetApp.getUi().alert("Higienização Concluída", relatorio, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

// ===================================================================
// ROTINAS DE ADMISSÃO E COBRANÇAS
// ===================================================================

function rotinaMatinalAgenteIntegrador() {
  var modoTeste = false;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return; }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAdmissao = ss.getSheetByName(CONFIG.SHEETS.ADMISSAO);
    var sheetHistorico = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);
    var sheetFiliais = ss.getSheetByName(CONFIG.SHEETS.FILIAIS);

    if (!sheetAdmissao || !sheetHistorico || !sheetFiliais) return;

    var hoje = new Date();
    var diasParaAdicionar = (hoje.getDay() === 5) ? 3 : (hoje.getDay() === 6) ? 2 : 1;
    var dataAlvo = new Date();
    dataAlvo.setDate(hoje.getDate() + diasParaAdicionar);
    var dataBuscaBR = Utilities.formatDate(dataAlvo, "GMT-3", "dd/MM/yyyy");

    var dadosAdmissao = sheetAdmissao.getDataRange().getValues();
    var dadosFiliais = sheetFiliais.getDataRange().getValues();

    var conjuntoIDs = new Set();
    var conjuntoNomesEData = new Set();

    if (sheetHistorico.getLastRow() > 1) {
      var dHist = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, 9).getValues();
      dHist.forEach(function(l) {
        if (l[0]) conjuntoIDs.add(normalizarID(l[0]));
        if (l[1] && l[3]) conjuntoNomesEData.add(normalizarTexto(l[1]) + "|" + formatarDataBR(l[3]));
      });
    }

    var mapaEmails = {};
    for (var j = 1; j < dadosFiliais.length; j++) {
      var fil = normalizarFilialId(dadosFiliais[j][0]);
      if (fil) {
        var emails = [dadosFiliais[j][3], dadosFiliais[j][4], dadosFiliais[j][5], dadosFiliais[j][6], dadosFiliais[j][7]].filter(Boolean);
        mapaEmails[fil] = emails.join(",");
      }
    }

    var novasLinhas = [];
    for (var i = 1; i < dadosAdmissao.length; i++) {
      var idOriginal = String(dadosAdmissao[i][0]).trim();
      if (!idOriginal) continue;
      var idLimpo = normalizarID(idOriginal);

      var nome = dadosAdmissao[i][1];
      var cargo = normalizarTexto(dadosAdmissao[i][2]);
      var filial = normalizarFilialId(dadosAdmissao[i][4]);
      var dataBR = formatarDataBR(dadosAdmissao[i][7]);

      var chave = normalizarTexto(nome) + "|" + dataBR;
      var cargoValido = CONFIG.CARGOS_PILOTO.some(function(c) { return cargo.includes(normalizarTexto(c)); });

      if (cargoValido && !conjuntoIDs.has(idLimpo) && !conjuntoNomesEData.has(chave) && dataBR === dataBuscaBR) {
        var emailDest = mapaEmails[filial] || "";
        if (emailDest) {
          var dest = modoTeste ? Session.getActiveUser().getEmail() : emailDest;
          var assunto = (modoTeste ? "[TESTE] " : "") + "🎉 Novo Colaborador (Filial " + filial + "): " + nome;
          var html = montarEmailAdmissao(nome, dataBR);
          try { GmailApp.sendEmail(dest, assunto, "", { htmlBody: html }); } catch(e){}
        }

        novasLinhas.push([idOriginal, nome, filial, dataBR, "PENDENTE", "PENDENTE", "PENDENTE", "EM ACOMPANHAMENTO", 0]);
        conjuntoIDs.add(idLimpo);
        conjuntoNomesEData.add(chave);
      }
    }

    if (novasLinhas.length > 0) {
      sheetHistorico.getRange(sheetHistorico.getLastRow() + 1, 1, novasLinhas.length, 9).setValues(novasLinhas);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

function rotinaNotificacaoFeedbacks() {
  processarDesligamentos();
  sincronizarFeedbacksUniversidade();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetH = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);
  var sheetF = ss.getSheetByName(CONFIG.SHEETS.FILIAIS);
  if (!sheetH || sheetH.getLastRow() <= 1) return;

  var dadosH = sheetH.getRange(2, 1, sheetH.getLastRow() - 1, 9).getValues();
  var dadosF = sheetF.getDataRange().getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  var pendenciasPorFilial = {};

  for (var i = 0; i < dadosH.length; i++) {
    var statusProc = String(dadosH[i][7] || '').toUpperCase();

    // Pula Desligados, Concluídos e Pausados (Férias/Afastados)
    if (statusProc === "DESLIGADO" || statusProc === "CONCLUIDO" || statusProc === "PAUSADO") continue;

    var dtAdm = parseDataBR(dadosH[i][3]);
    if (!dtAdm) continue;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);
    if (diasCasa < 15) continue;

    var f15 = String(dadosH[i][4] || '').toUpperCase();
    var f30 = String(dadosH[i][5] || '').toUpperCase();
    var f45 = String(dadosH[i][6] || '').toUpperCase();

    var filial = normalizarFilialId(dadosH[i][2]);
    var atraso = Number(dadosH[i][8]) || 0;

    if (atraso > 0) {
      var emails = buscarEmailsFilial(filial, dadosF);
      if (emails) {
        if (!pendenciasPorFilial[filial]) pendenciasPorFilial[filial] = { emails: emails, itens: [] };
        pendenciasPorFilial[filial].itens.push({
          nome: dadosH[i][1],
          diasCasa: diasCasa,
          f15: f15, f30: f30, f45: f45,
          atraso: atraso
        });
      }
    }
  }

  for (var filialKey in pendenciasPorFilial) {
    var grupo = pendenciasPorFilial[filialKey];
    enviarAlertaFeedback(grupo.emails, filialKey, grupo.itens);
  }
}

function buscarEmailsFilial(filial, dadosF) {
  var filNorm = normalizarFilialId(filial);
  for (var j = 1; j < dadosF.length; j++) {
    if (normalizarFilialId(dadosF[j][0]) === filNorm) {
      var gte = dadosF[j][3] ? String(dadosF[j][3]).trim() : "";
      var lid = dadosF[j][6] ? String(dadosF[j][6]).trim() : "";
      if (gte && lid) return gte + "," + lid;
      if (gte) return gte;
      if (lid) return lid;
    }
  }
  return null;
}

function enviarAlertaFeedback(email, filial, itensPendentes) {
  var assunto = "⚠️ AÇÃO: " + itensPendentes.length + " Pendência(s) de Feedback - Filial " + filial;
  var html = montarEmailFeedback(filial, itensPendentes);
  try { GmailApp.sendEmail(email, assunto, "", { htmlBody: html }); } catch(e){}
}

function processarDesligamentos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAtivos = ss.getSheetByName(CONFIG.SHEETS.ATIVOS);
  var sheetHis = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);
  if (!sheetAtivos || !sheetHis || sheetHis.getLastRow() <= 1) return;

  var dadosAtivos = sheetAtivos.getDataRange().getValues();
  if (dadosAtivos.length <= 1) return;

  // Mapeamento dinâmico dos cabeçalhos da aba DADOS_ATIVOS
  var headerMap = mapearCabecalhos(dadosAtivos[0]);
  var colId = encontrarColuna(headerMap, ["id", "matricula", "codigo"]);
  var colSituacao = encontrarColuna(headerMap, ["situacao", "status", "situacao_colaborador"]);

  // Fallback baseado na estrutura verificada (Coluna B = ID [1], Coluna H = Situacao [7])
  if (colId === -1) colId = 1;
  if (colSituacao === -1) colSituacao = 7;

  var ativosSet = new Set();
  var pausadosSet = new Set();
  var desligadosSet = new Set();

  for (var i = 1; i < dadosAtivos.length; i++) {
    var id = normalizarID(dadosAtivos[i][colId]);
    var sit = normalizarTexto(dadosAtivos[i][colSituacao]);
    if (!id) continue;

    if (sit.includes("EM ATIVIDADE NORMAL")) {
      ativosSet.add(id);
    } else if (sit.includes("DEMITIDO")) {
      desligadosSet.add(id);
    } else {
      // Férias, Afastamentos, Licença, Aposentadoria, etc.
      pausadosSet.add(id);
    }
  }

  var rangeHis = sheetHis.getRange(2, 1, sheetHis.getLastRow() - 1, 9);
  var dadosHis = rangeHis.getValues();
  var alterou = false;

  for (var k = 0; k < dadosHis.length; k++) {
    var idHist = normalizarID(dadosHis[k][0]);
    if (!idHist) continue;

    var statusAtual = String(dadosHis[k][7] || '').toUpperCase();

    // Regra 1: Demitido ou Ausente da Planilha -> Marcar como DESLIGADO e zerar atraso
    if (desligadosSet.has(idHist) || (!ativosSet.has(idHist) && !pausadosSet.has(idHist))) {
      if (statusAtual !== "DESLIGADO") {
        dadosHis[k][7] = "DESLIGADO";
        dadosHis[k][8] = 0;
        alterou = true;
      }
    } 
    // Regra 2: Férias / Afastamentos -> Marcar como PAUSADO e zerar atraso/cobrança
    else if (pausadosSet.has(idHist)) {
      if (statusAtual !== "PAUSADO") {
        dadosHis[k][7] = "PAUSADO";
        dadosHis[k][8] = 0;
        alterou = true;
      }
    }
    // Regra 3: Em Atividade Normal -> Reativar se estava pausado ou desligado indevidamente
    else if (ativosSet.has(idHist) && (statusAtual === "DESLIGADO" || statusAtual === "PAUSADO")) {
      dadosHis[k][7] = "EM ACOMPANHAMENTO";
      alterou = true;
    }
  }

  if (alterou) {
    rangeHis.setValues(dadosHis);
    SpreadsheetApp.flush();
  }
}

function sincronizarFeedbacksUniversidade() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetF = ss.getSheetByName(CONFIG.SHEETS.FEEDBACK);
  var sheetH = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);
  if (!sheetF || !sheetH || sheetH.getLastRow() <= 1) return;

  var dadosF = sheetF.getDataRange().getValues();
  var mapaFeedbacks = {};
  for (var i = 1; i < dadosF.length; i++) {
    var id = normalizarID(dadosF[i][0]);
    if (!id) continue;
    mapaFeedbacks[id] = {
      f15: String(dadosF[i][8]) === "1" || normalizarTexto(dadosF[i][8]) === "OK",
      f30: String(dadosF[i][9]) === "1" || normalizarTexto(dadosF[i][9]) === "OK",
      f45: String(dadosF[i][10]) === "1" || normalizarTexto(dadosF[i][10]) === "OK"
    };
  }

  var rangeH = sheetH.getRange(2, 1, sheetH.getLastRow() - 1, 9);
  var dadosH = rangeH.getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  for (var k = 0; k < dadosH.length; k++) {
    var idHist = normalizarID(dadosH[k][0]);
    var statusAtual = String(dadosH[k][7] || '').toUpperCase();

    // Se estiver DESLIGADO ou PAUSADO, mantemos o atraso zerado e pulamos recálculo
    if (statusAtual === "DESLIGADO" || statusAtual === "PAUSADO") {
      dadosH[k][8] = 0;
      continue;
    }

    if (mapaFeedbacks[idHist]) {
      var fb = mapaFeedbacks[idHist];
      if (fb.f15) dadosH[k][4] = "OK";
      if (fb.f30) dadosH[k][5] = "OK";
      if (fb.f45) dadosH[k][6] = "OK";
    }

    var f15 = String(dadosH[k][4] || '').toUpperCase();
    var f30 = String(dadosH[k][5] || '').toUpperCase();
    var f45 = String(dadosH[k][6] || '').toUpperCase();

    if (f15 === "OK" && f30 === "OK" && f45 === "OK") {
      dadosH[k][7] = "CONCLUIDO";
      dadosH[k][8] = 0;
    } else {
      dadosH[k][7] = "EM ACOMPANHAMENTO";
      var dtAdm = parseDataBR(dadosH[k][3]);
      var diasAtraso = 0;
      if (dtAdm) {
        var diasCasa = calcularDiasCasa(dtAdm, hoje);
        if (diasCasa >= 15 && f15 !== "OK") {
          diasAtraso = diasCasa - 15;
        } else if (diasCasa >= 30 && f30 !== "OK") {
          diasAtraso = diasCasa - 30;
        } else if (diasCasa >= 45 && f45 !== "OK") {
          diasAtraso = diasCasa - 45;
        }
      }
      dadosH[k][8] = diasAtraso;
    }
  }
  rangeH.setValues(dadosH);
  SpreadsheetApp.flush();
}

function enviarRelatorioSemanalCoordenadores() {
  sincronizarFeedbacksUniversidade();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetH = ss.getSheetByName(CONFIG.SHEETS.HISTORICO_NOVO);
  var sheetF = ss.getSheetByName(CONFIG.SHEETS.FILIAIS);
  if (!sheetH || sheetH.getLastRow() <= 1) return;

  var dadosH = sheetH.getRange(2, 1, sheetH.getLastRow() - 1, 9).getValues();
  var dadosF = sheetF.getDataRange().getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  var coordMap = {};
  for(var j = 1; j < dadosF.length; j++) {
    var fil = normalizarFilialId(dadosF[j][0]);
    if (fil) {
      coordMap[fil] = { email: dadosF[j][4], regional: dadosF[j][2], gerenteGP: dadosF[j][5], emailRegional: dadosF[j][7] };
    }
  }

  var relatorios = {};
  for (var i = 0; i < dadosH.length; i++) {
    var statusProc = String(dadosH[i][7] || '').toUpperCase();
    if (statusProc === "DESLIGADO" || statusProc === "CONCLUIDO" || statusProc === "PAUSADO") continue;

    var fil = normalizarFilialId(dadosH[i][2]);
    var info = coordMap[fil];
    if (!info || !info.email) continue;

    var dtAdm = parseDataBR(dadosH[i][3]);
    if (!dtAdm) continue;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);
    if (diasCasa < 15) continue;

    var pend = [];
    if (diasCasa >= 15 && dadosH[i][4] !== "OK") pend.push("Feedback 15 dias");
    if (diasCasa >= 30 && dadosH[i][5] !== "OK") pend.push("Feedback 30 dias");
    if (diasCasa >= 45 && dadosH[i][6] !== "OK") pend.push("Feedback 45 dias");

    if (pend.length > 0) {
      var dest = info.email.trim();
      if (!relatorios[dest]) relatorios[dest] = [];
      relatorios[dest].push({ filial: fil, regional: info.regional || "Sem Regional", nome: dadosH[i][1], detalhe: pend.join(", "), gerenteGP: info.gerenteGP, emailRegional: info.emailRegional });
    }
  }

  for (var email in relatorios) {
    var grupos = {};
    relatorios[email].forEach(function(r) { (grupos[r.regional] = grupos[r.regional] || []).push(r); });
    var html = montarEmailRelatorioSemanal(grupos, relatorios[email].length);
    var cc = Array.from(new Set(relatorios[email].map(function(x){ return x.gerenteGP || x.emailRegional; }))).filter(Boolean).join(",");
    var opt = { htmlBody: html };
    if (cc) opt.cc = cc;
    try { GmailApp.sendEmail(email, "📊 Relatório Semanal de Pendências de Feedback", "", opt); } catch(e){}
  }
}

// ===================================================================
// TEMPLATES HTML DE E-MAIL
// ===================================================================

function montarEmailAdmissao(nomeColaborador, dataAdmFormatada) {
  return '' +
'<div style="font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f8; padding: 24px;">' +
  '<div style="background: linear-gradient(135deg, #0056b3, #007bff); border-radius: 10px 10px 0 0; padding: 28px 24px; text-align: center;">' +
    '<div style="font-size: 34px; margin-bottom: 6px;">🎉</div>' +
    '<h1 style="color: #ffffff; margin: 0; font-size: 20px; font-weight: 600;">Novo Colaborador a Bordo!</h1>' +
  '</div>' +
  '<div style="background-color: #ffffff; padding: 28px 24px; border-radius: 0 0 10px 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);">' +
    '<p style="color: #333333; font-size: 15px; margin-top: 0;">Olá, gerente!</p>' +
    '<p style="color: #333333; font-size: 15px; line-height: 1.6;">É com entusiasmo que confirmamos a chegada de mais um talento para o time.</p>' +
    '<div style="background-color: #f0f7ff; border-left: 4px solid #007bff; border-radius: 6px; padding: 16px 18px; margin: 20px 0;">' +
      '<p style="margin: 0 0 6px 0; font-size: 13px; color: #666666; text-transform: uppercase; letter-spacing: 0.5px;">Colaborador</p>' +
      '<p style="margin: 0 0 12px 0; font-size: 18px; color: #0056b3; font-weight: 700;">' + nomeColaborador + '</p>' +
      '<p style="margin: 0 0 6px 0; font-size: 13px; color: #666666; text-transform: uppercase; letter-spacing: 0.5px;">Início</p>' +
      '<p style="margin: 0; font-size: 16px; color: #333333; font-weight: 600;">📅 ' + dataAdmFormatada + '</p>' +
    '</div>' +
    '<p style="color: #333333; font-size: 15px; line-height: 1.6;">Para garantir que a integração seja um sucesso, preparamos um roteiro com as diretrizes essenciais:</p>' +
    '<div style="margin: 20px 0;">' +
      '<a href="https://docs.google.com/presentation/d/1TObH1pVrxi5XCye364Esr4Rfz9x1jIcHzlc7TSFxngg/edit?usp=sharing" style="display:block; background-color:#0056b3; color:#ffffff; text-decoration:none; padding:12px 18px; border-radius:6px; font-size:14px; font-weight:600; margin-bottom:10px; text-align:center;">📘 Acessar E-Book Gerente</a>' +
      '<a href="https://gamma.app/docs/E-BOOK-Integracao-de-Novos-Colaboradores-Gerentes-ftz6vrsw5y0bmxi?mode=doc" style="display:block; background-color:#ffffff; color:#0056b3; border: 1px solid #0056b3; text-decoration:none; padding:11px 18px; border-radius:6px; font-size:14px; font-weight:600; margin-bottom:10px; text-align:center;">📱 E-Book Digital Gerente</a>' +
      '<a href="https://docs.google.com/presentation/d/1nVs14MjkLG5oaQrm5Zd9Wno-hLr2x1PlRLtXVWY3Y1Y/edit?usp=sharing" style="display:block; background-color:#ffffff; color:#0056b3; border: 1px solid #0056b3; text-decoration:none; padding:11px 18px; border-radius:6px; font-size:14px; font-weight:600; text-align:center;">👤 E-Book Colaborador (envie ao novo membro)</a>' +
    '</div>' +
    '<p style="color: #666666; font-size: 13px; margin-top: 24px; border-top: 1px solid #eeeeee; padding-top: 16px;">Atenciosamente,<br><strong style="color:#333333;">Gestão de Pessoas</strong></p>' +
  '</div>' +
'</div>';
}

function montarEmailFeedback(filial, itensPendentes) {
  var montarBadges = function(diasCasa, f15, f30, f45) {
    var etapas = [
      { label: "15 dias", limite: 15, statusOk: f15 === "OK" },
      { label: "30 dias", limite: 30, statusOk: f30 === "OK" },
      { label: "45 dias", limite: 45, statusOk: f45 === "OK" }
    ];

    return etapas.map(function(e) {
      if (diasCasa < e.limite) {
        return '<span style="display:inline-block; background-color:#f1f1f1; color:#999999; border-radius:20px; padding:5px 12px; font-size:11px; font-weight:600; margin:0 5px 5px 0;">◽ ' + e.label + '</span>';
      }
      if (e.statusOk) {
        return '<span style="display:inline-block; background-color:#e6f4ea; color:#1e7e34; border-radius:20px; padding:5px 12px; font-size:11px; font-weight:600; margin:0 5px 5px 0;">✅ ' + e.label + '</span>';
      }
      return '<span style="display:inline-block; background-color:#fdecea; color:#c0392b; border-radius:20px; padding:5px 12px; font-size:11px; font-weight:600; margin:0 5px 5px 0;">⏳ ' + e.label + '</span>';
    }).join("");
  };

  var itensOrdenados = itensPendentes.slice().sort(function(a, b) { return b.atraso - a.atraso; });

  var cardsHtml = itensOrdenados.map(function(item) {
    var badgesHtml = montarBadges(item.diasCasa, item.f15, item.f30, item.f45);
    return '' +
'<div style="background-color: #fafafa; border-radius: 8px; padding: 14px 16px; margin: 0 0 12px 0;">' +
  '<div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:8px;">' +
    '<p style="margin:0; font-size:15px; color:#333333; font-weight:700;">' + item.nome + '</p>' +
    '<span style="white-space:nowrap; font-size:12px; color:#856404; background-color:#fff3cd; padding:3px 10px; border-radius:12px; font-weight:600;">+' + item.atraso + ' dia(s)</span>' +
  '</div>' +
  '<p style="margin:0 0 8px 0; font-size:12px; color:#888888;">' + item.diasCasa + ' dias de casa</p>' +
  '<div>' + badgesHtml + '</div>' +
'</div>';
  }).join("");

  var totalAtraso = itensOrdenados.reduce(function(soma, item) { return soma + item.atraso; }, 0);

  return '' +
'<div style="font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f4f6f8; padding: 24px;">' +
  '<div style="background: linear-gradient(135deg, #d9534f, #e8746f); border-radius: 10px 10px 0 0; padding: 24px; text-align: center;">' +
    '<div style="font-size: 30px; margin-bottom: 4px;">⚠️</div>' +
    '<h1 style="color: #ffffff; margin: 0 0 6px 0; font-size: 19px; font-weight: 600;">Pendências de Feedback</h1>' +
    '<p style="color:#fbdcda; margin:0; font-size:13px;">Filial ' + filial + ' · ' + itensOrdenados.length + ' colaborador(es) pendente(s)</p>' +
  '</div>' +
  '<div style="background-color: #ffffff; padding: 26px 24px; border-radius: 0 0 10px 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);">' +
    '<p style="color: #333333; font-size: 15px; margin-top: 0;">Olá, gerente!</p>' +
    '<p style="color: #333333; font-size: 14px; line-height: 1.6;">Os colaboradores abaixo, da <strong>Filial ' + filial + '</strong>, estão com feedback pendente:</p>' +
    cardsHtml +
    '<div style="background-color: #fff3cd; border-radius: 8px; padding: 14px 16px; margin: 6px 0 18px 0; text-align:center;">' +
      '<p style="margin:0; font-size: 14px; color: #856404;">Atraso total acumulado da filial: <strong style="font-size:16px;">' + totalAtraso + ' dias</strong></p>' +
    '</div>' +
    '<div style="background-color: #e8f0fe; border-left: 4px solid #4285f4; border-radius: 6px; padding: 14px 16px; margin: 18px 0;">' +
      '<p style="margin:0; font-size: 13px; color: #1a4a8a; line-height:1.5;">ℹ️ Após a conversa, <strong>o próprio colaborador</strong> deve registrar o feedback recebido na Universidade Luiza.</p>' +
    '</div>' +
    '<a href="https://universidadeluiza.com.br/app/menu/acompanhamento-do-periodo-de-experiencia-novos-colaboradores-de-loja" style="display:block; background-color:#0056b3; color:#ffffff; text-decoration:none; padding:12px 18px; border-radius:6px; font-size:14px; font-weight:600; text-align:center; margin: 18px 0 10px 0;">📝 Registrar Feedback na Universidade Luiza</a>' +
    '<a href="https://docs.google.com/presentation/d/1TObH1pVrxi5XCye364Esr4Rfz9x1jIcHzlc7TSFxngg/edit?usp=sharing" style="display:block; color:#0056b3; text-decoration:underline; font-size:13px; text-align:center;">📘 Guia de Feedback</a>' +
    '<p style="color: #999999; font-size: 12px; margin-top: 24px; border-top: 1px solid #eeeeee; padding-top: 16px; text-align:center;">Gestão de Pessoas</p>' +
  '</div>' +
'</div>';
}

function montarEmailRelatorioSemanal(gruposPorRegional, totalPendencias) {
  var blocosRegionais = "";
  for (var regionalName in gruposPorRegional) {
    var itens = gruposPorRegional[regionalName];
    var linhasHtml = itens.map(function(item) {
      return '' +
'<div style="display:flex; justify-content:space-between; align-items:flex-start; padding:10px 0; border-bottom:1px solid #f0f0f0;">' +
  '<div>' +
    '<p style="margin:0; font-size:14px; color:#333333; font-weight:600;">Filial ' + item.filial + ' — ' + item.nome + '</p>' +
    '<p style="margin:2px 0 0 0; font-size:12px; color:#c0392b;">⚠️ ' + item.detalhe + '</p>' +
  '</div>' +
'</div>';
    }).join("");

    blocosRegionais += '' +
'<div style="margin-bottom:18px; border:1px solid #e6e6e6; border-radius:8px; background-color:#fcfcfc; overflow:hidden;">' +
  '<div style="background-color:#fbe9e7; padding:10px 16px;">' +
    '<h3 style="margin:0; color:#c0392b; font-size:14px; font-weight:700;">📍 Regional: ' + regionalName + '</h3>' +
  '</div>' +
  '<div style="padding:6px 16px;">' + linhasHtml + '</div>' +
'</div>';
  }

  return '' +
'<div style="font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif; max-width: 640px; margin: 0 auto; background-color: #f4f6f8; padding: 24px;">' +
  '<div style="background: linear-gradient(135deg, #0056b3, #003d82); border-radius: 10px 10px 0 0; padding: 26px 24px; text-align:center;">' +
    '<div style="font-size:30px; margin-bottom:4px;">📊</div>' +
    '<h1 style="color:#ffffff; margin:0 0 6px 0; font-size:19px; font-weight:600;">Relatório Semanal de Feedbacks</h1>' +
    '<p style="color:#cfe0f5; margin:0; font-size:13px;">' + totalPendencias + ' pendência(s) identificada(s)</p>' +
  '</div>' +
  '<div style="background-color:#ffffff; padding:24px; border-radius:0 0 10px 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.06);">' +
    '<p style="color:#333333; font-size:14px; margin-top:0;">Olá,</p>' +
    '<p style="color:#333333; font-size:14px; line-height:1.6;">Abaixo estão as pendências de feedback de suas filiais, agrupadas por regional:</p>' +
    blocosRegionais +
    '<p style="color:#999999; font-size:12px; margin-top:20px; border-top:1px solid #eeeeee; padding-top:14px; text-align:center;">Este é um e-mail automático do Sistema de GP Lojas Magalu.<br><strong>Gestão de Pessoas</strong></p>' +
  '</div>' +
'</div>';
}

// ===================================================================
// UTILITÁRIOS E MIDDLEWARE DE FILIAL
// ===================================================================

/**
 * Middleware Global Obrigatório para Higienização e Normalização de Filiais
 * Regra: Converte para número e subtrai 3000 caso seja maior que 3000 (Séries 3000/4000)
 */
function normalizarFilialId(val) {
  if (val === undefined || val === null || val === "") return "";
  var strLimpa = String(val).replace(/\D/g, '');
  if (!strLimpa) return String(val).trim();
  var num = parseInt(strLimpa, 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) {
    num -= 3000;
  }
  return String(num);
}

function normalizarTexto(t) {
  return t ? t.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim() : "";
}

function normalizarID(id) {
  if (id === undefined || id === null) return "";
  var str = String(id).trim().replace(/^0+/, '');
  return str === "" ? "0" : str;
}

function formatarDataBR(valor) {
  if (!valor) return "";
  if (valor instanceof Date) return Utilities.formatDate(valor, "GMT-3", "dd/MM/yyyy");
  var str = valor.toString().trim();
  if (str.indexOf("-") > -1) {
    var p = str.split("-");
    if (p.length === 3) return p[2].substring(0,2) + "/" + p[1] + "/" + p[0];
  }
  return str;
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
  const map = {};
  headerRow.forEach((col, idx) => { map[normalizarTexto(col).replace(/[\s_]+/g, '')] = idx; });
  return map;
}

function encontrarColuna(headerMap, candidatos) {
  for (let cand of candidatos) {
    const clean = normalizarTexto(cand).replace(/[\s_]+/g, '');
    for (let key in headerMap) {
      if (key.includes(clean)) return headerMap[key];
    }
  }
  return -1;
}

function limparCacheDashboard() {
  CacheService.getScriptCache().remove(CONFIG.CACHE_KEY);
  console.log("Cache limpo.");
}

function configurarTodosGatilhos() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) ScriptApp.deleteTrigger(ts[i]);

  ScriptApp.newTrigger('rotinaMatinalAgenteIntegrador').timeBased().atHour(10).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('sincronizarFeedbacksUniversidade').timeBased().atHour(12).everyDays(1).create();
  ScriptApp.newTrigger('rotinaNotificacaoFeedbacks').timeBased().atHour(13).everyDays(1).create();
  ScriptApp.newTrigger('enviarRelatorioSemanalCoordenadores').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}
