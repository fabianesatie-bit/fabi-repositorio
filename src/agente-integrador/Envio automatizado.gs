/**
 * SISTEMA GP MAGALU - GESTÃO INTEGRADA (ADMISSÃO + FEEDBACKS)
 * Versão 6.4 - Cobranças de feedback agrupadas por filial
 *
 * O QUE MUDOU EM RELAÇÃO À v6.3:
 * 5. [AGRUPAMENTO] rotinaNotificacaoFeedbacks() não envia mais um e-mail por colaborador pendente.
 *    Agora acumula todas as pendências em um objeto agrupado por filial (pendenciasPorFilial) e
 *    dispara UM único e-mail consolidado por filial, listando todos os colaboradores pendentes
 *    ordenados por maior atraso. Isso reduz drasticamente o volume de e-mails para gerentes de
 *    filiais com vários colaboradores pendentes ao mesmo tempo.
 * 6. [FUNÇÕES ATUALIZADAS] enviarAlertaFeedback() e montarEmailFeedback() agora recebem uma lista
 *    de colaboradores em vez de dados de um único colaborador.
 *
 * O QUE MUDOU EM RELAÇÃO À v6.2:
 * 1. [BUG CRÍTICO] Os Sets de controle anti-duplicidade (conjuntoIDs / conjuntoNomesEData) agora são
 *    atualizados DURANTE o próprio loop de leitura. Antes, duas linhas duplicadas na mesma planilha de
 *    origem (mesmo ID ou mesmo Nome+Data) passavam juntas pela validação, pois ela só comparava com o
 *    que já existia no Histórico ANTES da execução começar.
 * 2. [ROBUSTEZ] Adicionado LockService.getScriptLock() em prepararListaModelo() para evitar que duas
 *    execuções simultâneas (ex: gatilho automático + execução manual) leiam o mesmo snapshot do
 *    Histórico e gravem duplicado.
 * 3. [CONSISTÊNCIA] O ID agora é normalizado (removendo zeros à esquerda) antes de entrar no Set,
 *    igual já era feito com o campo Filial - evita falha de correspondência entre "123" e "0123".
 * 4. [VISUAL] Todos os e-mails (admissão, feedback e relatório semanal) foram redesenhados com um
 *    layout mais bonito: cabeçalho colorido, cards, ícones e melhor hierarquia visual.
 *
 * Nenhuma função foi removida. Todas as rotinas e gatilhos já em produção continuam com o mesmo nome
 * e mesmo comportamento funcional (apenas corrigidos/embelezados).
 */

// --- BLOCO 1: ROTINA DE ADMISSÃO ---

function rotinaMatinalAgenteIntegrador() {
  console.log("Iniciando rotina de admissão...");
  // Altere para true se quiser apenas testar os dados sem disparar e-mails para os gerentes reais
  var modoModoTeste = false;

  var encontrouRegistros = prepararListaModelo();
  if (encontrouRegistros) {
    dispararEnvioMensagemAdmissao(modoModoTeste);
  } else {
    console.log("Nenhuma admissão válida encontrada para a data alvo.");
  }
}

function prepararListaModelo() {
  // Trava de concorrência: impede que duas execuções simultâneas leiam o mesmo
  // snapshot do Histórico e acabem gravando os mesmos registros duas vezes.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // espera até 30s por uma execução concorrente liberar o lock
  } catch (e) {
    console.log("Não foi possível obter o lock (outra execução em andamento). Abortando com segurança.");
    return false;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAdmissao = ss.getSheetByName("Atualizacao_admissao");
    var sheetModelo = ss.getSheetByName("MODELO_AGENTE_INTEGRADOR");
    var sheetFiliais = ss.getSheetByName("Base filiais");
    var sheetHistorico = ss.getSheetByName("Historico");

    var cargosPiloto = ["VENDEDOR", "ASSISTENTE VENDAS SR", "ASSISTENTE VENDAS JR", "AUX LIMPEZA", "LIDER COMERCIAL", "LIDER DE ESTOQUE"];

    var hoje = new Date();
    var diasParaAdicionar = (hoje.getDay() === 5) ? 3 : (hoje.getDay() === 6) ? 2 : 1;
    var dataAlvo = new Date();
    dataAlvo.setDate(hoje.getDate() + diasParaAdicionar);

    var dataBuscaBR = Utilities.formatDate(dataAlvo, "GMT-3", "dd/MM/yyyy");

    var dadosAdmissao = sheetAdmissao.getDataRange().getValues();
    var dadosFiliais = sheetFiliais.getDataRange().getValues();

    // Bloco de validação inteligente anti-duplicidade (ID e Nome+Data)
    var conjuntoIDs = new Set();
    var conjuntoNomesEData = new Set();

    if (sheetHistorico.getLastRow() > 1) {
      var dadosHistoricoCompleto = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, sheetHistorico.getLastColumn()).getValues();

      dadosHistoricoCompleto.forEach(function(linha) {
        var idHist = normalizarID(linha[3]);
        var nomeHist = normalizarTexto(linha[5]);
        var dataHist = String(linha[6]).trim();

        if (idHist) conjuntoIDs.add(idHist);
        if (nomeHist && dataHist) {
          conjuntoNomesEData.add(nomeHist + "|" + dataHist);
        }
      });
    }

    var novasLinhas = [];
    for (var i = 1; i < dadosAdmissao.length; i++) {
      var idFuncOriginal = String(dadosAdmissao[i][0]).trim();
      if (!idFuncOriginal) continue;
      var idFunc = normalizarID(idFuncOriginal);

      var nomeBusca = normalizarTexto(dadosAdmissao[i][1]);
      var cargoBusca = normalizarTexto(dadosAdmissao[i][2]);

      var valorDataBruto = dadosAdmissao[i][7];
      var dataFormatadaBR = "";

      if (valorDataBruto instanceof Date) {
        dataFormatadaBR = Utilities.formatDate(valorDataBruto, "GMT-3", "dd/MM/yyyy");
      } else if (valorDataBruto && valorDataBruto.toString().indexOf("-") > -1) {
        var partes = valorDataBruto.toString().split("-");
        dataFormatadaBR = partes[2].substring(0,2) + "/" + partes[1] + "/" + partes[0];
      } else {
        dataFormatadaBR = valorDataBruto.toString().trim();
      }

      var datasCoincidem = (dataFormatadaBR === dataBuscaBR);

      var chaveNomeData = nomeBusca + "|" + dataFormatadaBR;
      var idJaExiste = conjuntoIDs.has(idFunc);
      var nomeEDataJaExiste = conjuntoNomesEData.has(chaveNomeData);

      if (cargosPiloto.indexOf(cargoBusca) !== -1 && !idJaExiste && !nomeEDataJaExiste && datasCoincidem) {
        var filialBusca = String(dadosAdmissao[i][4]).trim().replace(/^0+/, '');
        var emailsDestino = "";

        for (var j = 1; j < dadosFiliais.length; j++) {
          if (String(dadosFiliais[j][0]).trim().replace(/^0+/, '') === filialBusca) {
            var listaEmails = [];
            if (dadosFiliais[j][3]) listaEmails.push(String(dadosFiliais[j][3]).trim());
            if (dadosFiliais[j][4]) listaEmails.push(String(dadosFiliais[j][4]).trim());
            if (dadosFiliais[j][5]) listaEmails.push(String(dadosFiliais[j][5]).trim());
            if (dadosFiliais[j][6]) listaEmails.push(String(dadosFiliais[j][6]).trim());
            if (dadosFiliais[j][7]) listaEmails.push(String(dadosFiliais[j][7]).trim());

            // Alterado para vírgula (,), garantindo o padrão do GmailApp
            emailsDestino = listaEmails.join(",");
            break;
          }
        }
        novasLinhas.push([dadosAdmissao[i][4], dadosAdmissao[i][6], dadosAdmissao[i][5], idFuncOriginal, "", dadosAdmissao[i][1], dataFormatadaBR, "Nova Admissão", emailsDestino, "SIM", "AGENDADO", "", "", "", "", "0"]);

        // *** CORREÇÃO PRINCIPAL DO BUG DE DUPLICIDADE ***
        // Assim que a linha é aceita, ela entra imediatamente nos Sets de controle.
        // Isso garante que, se houver outra linha duplicada mais abaixo na MESMA planilha
        // de origem (mesmo ID ou mesmo Nome+Data), ela será bloqueada nesta mesma execução.
        conjuntoIDs.add(idFunc);
        conjuntoNomesEData.add(chaveNomeData);
      }
    }

    if (novasLinhas.length > 0) {
      if (sheetModelo.getLastRow() > 1) {
        sheetModelo.getRange(2, 1, sheetModelo.getLastRow() - 1, sheetModelo.getLastColumn()).clearContent();
      }
      sheetModelo.getRange(2, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);
      sheetHistorico.getRange(sheetHistorico.getLastRow() + 1, 1, novasLinhas.length, novasLinhas[0].length).setValues(novasLinhas);

      SpreadsheetApp.flush();
      return true;
    }
    return false;

  } finally {
    lock.releaseLock();
  }
}

function dispararEnvioMensagemAdmissao(modoModoTeste) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("MODELO_AGENTE_INTEGRADOR");
  var dados = sheet.getDataRange().getValues();
  var emailSuporteTeste = Session.getActiveUser().getEmail();

  var ultLinha = sheet.getLastRow();
  if (ultLinha <= 1) return;

  // Carrega a coluna de status (Coluna K - índice 11) para gravação em lote
  var rangeStatus = sheet.getRange(2, 11, ultLinha - 1, 1);
  var statusValues = rangeStatus.getValues();

  for (var i = 1; i < dados.length; i++) {
    var statusAtual = String(dados[i][10]).trim().toUpperCase();
    var emailDest = String(dados[i][8]).trim();

    // Corrige dinamicamente possíveis ponto e vírgulas remanescentes da base de dados
    if (emailDest.indexOf(";") > -1) {
      emailDest = emailDest.replace(/;/g, ",");
    }

    if (statusAtual === "AGENDADO" && emailDest !== "") {
      var destinatarioFinal = modoModoTeste ? emailSuporteTeste : emailDest;
      var assunto = (modoModoTeste ? "[TESTE] " : "") + "🎉 Novo Colaborador (Filial " + dados[i][0] + "): " + dados[i][5];

      var dataAdmFormatada = dados[i][6];
      if (dataAdmFormatada instanceof Date) {
        dataAdmFormatada = Utilities.formatDate(dataAdmFormatada, "GMT-3", "dd/MM/yyyy");
      } else {
        dataAdmFormatada = String(dataAdmFormatada).trim();
      }

      var corpoHtml = montarEmailAdmissao(dados[i][5], dataAdmFormatada);

      try {
        GmailApp.sendEmail(destinatarioFinal, assunto, "", { htmlBody: corpoHtml });
        statusValues[i - 1][0] = modoModoTeste ? "TESTADO OK" : "ENVIADO";
      } catch(e) {
        console.log("Erro ao enviar e-mail para " + destinatarioFinal + ": " + e.message);
        statusValues[i - 1][0] = "ERRO";
      }
    }
  }

  // Grava todos os status atualizados de uma só vez na planilha (Previne o timeout)
  rangeStatus.setValues(statusValues);
  SpreadsheetApp.flush();
}

// --- BLOCO 2: ROTINA DE FEEDBACKS ---

function rotinaNotificacaoFeedbacks() {
  processarDesligamentos();
  sincronizarFeedbacksUniversidade();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetHistorico = ss.getSheetByName("Historico");
  var sheetFiliais = ss.getSheetByName("Base filiais");
  var dadosH = sheetHistorico.getDataRange().getValues();
  var dadosF = sheetFiliais.getDataRange().getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  // Agrupador: em vez de disparar um e-mail por colaborador, acumulamos todas as
  // pendências por filial e enviamos UM único e-mail consolidado por filial no final.
  var pendenciasPorFilial = {};

  for (var i = 1; i < dadosH.length; i++) {
    var numCols = dadosH[i].length;

    var statusColaborador = (numCols >= 18 && dadosH[i][17]) ? String(dadosH[i][17]).trim().toUpperCase() : "";
    if (statusColaborador === "DESLIGADO") {
      sheetHistorico.getRange(i + 1, 16).setValue(0);
      continue;
    }

    var dtAdm = parseDataBR(dadosH[i][6]);
    if (!dtAdm || isNaN(dtAdm.getTime())) continue;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);

    if (diasCasa < 15) {
      sheetHistorico.getRange(i + 1, 16).setValue(0);
      continue;
    }

    var f15 = (numCols >= 13 && dadosH[i][12]) ? String(dadosH[i][12]).trim().toUpperCase() : "";
    var f30 = (numCols >= 14 && dadosH[i][13]) ? String(dadosH[i][13]).trim().toUpperCase() : "";
    var f45 = (numCols >= 15 && dadosH[i][14]) ? String(dadosH[i][14]).trim().toUpperCase() : "";

    var nome = dadosH[i][5];
    var filial = String(dadosH[i][0]).trim().replace(/^0+/, '');

    var atraso = 0;
    if (diasCasa >= 15 && f15 !== "OK") {
      atraso = diasCasa - 15;
    } else if (diasCasa >= 30 && f30 !== "OK") {
      atraso = diasCasa - 30;
    } else if (diasCasa >= 45 && f45 !== "OK") {
      atraso = diasCasa - 45;
    }

    sheetHistorico.getRange(i + 1, 16).setValue(atraso);

    var deveEnviar = false;
    if (diasCasa >= 15 && f15 !== "OK") {
      deveEnviar = true;
    } else if (diasCasa >= 30 && f30 !== "OK") {
      deveEnviar = true;
    } else if (diasCasa >= 45 && f45 !== "OK") {
      deveEnviar = true;
    }

    if (deveEnviar) {
       var emails = buscarEmailsFilial(filial, dadosF);
       if (emails) {
         if (!pendenciasPorFilial[filial]) {
           pendenciasPorFilial[filial] = { emails: emails, itens: [] };
         }
         pendenciasPorFilial[filial].itens.push({
           nome: nome,
           diasCasa: diasCasa,
           f15: f15,
           f30: f30,
           f45: f45,
           atraso: atraso
         });
       }
    }
  }

  // Envia um único e-mail por filial, com todos os colaboradores pendentes agrupados.
  for (var filialKey in pendenciasPorFilial) {
    var grupo = pendenciasPorFilial[filialKey];
    enviarAlertaFeedback(grupo.emails, filialKey, grupo.itens);
  }
}

function buscarEmailsFilial(filial, dadosF) {
  for (var j = 1; j < dadosF.length; j++) {
    if (String(dadosF[j][0]).trim().replace(/^0+/, '') === filial) {
      var emailGerente = dadosF[j][3] ? String(dadosF[j][3]).trim() : "";
      var emailLider = dadosF[j][6] ? String(dadosF[j][6]).trim() : "";

      if (emailGerente && emailLider) {
        return emailGerente + "," + emailLider;
      } else if (emailGerente) {
        return emailGerente;
      } else if (emailLider) {
        return emailLider;
      }
    }
  }
  return null;
}

function enviarAlertaFeedback(email, filial, itensPendentes) {
  var assunto = "⚠️ AÇÃO: " + itensPendentes.length + " Pendência(s) de Feedback - Filial " + filial;

  var html = montarEmailFeedback(filial, itensPendentes);

  try {
    GmailApp.sendEmail(email, assunto, "", { htmlBody: html });
  } catch(e) {
    console.log("Erro ao enviar e-mail de feedback: " + e.message);
  }
}

function processarDesligamentos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDes = ss.getSheetByName("Atualizacao_desligamento");
  var sheetHis = ss.getSheetByName("Historico");
  if (!sheetDes || !sheetHis) return;

  var dadosDes = sheetDes.getDataRange().getValues();
  var idsDesligados = new Set(dadosDes.slice(1).map(function(r){ return String(r[0]).trim(); }));
  var dadosHis = sheetHis.getRange(2, 4, sheetHis.getLastRow() - 1, 1).getValues().flat().map(String);

  for (var i = 1; i < dadosHis.length + 1; i++) {
    var idHistorico = String(dadosHis[i-1]).trim();
    if (idHistorico && idsDesligados.has(idHistorico)) {
      sheetHis.getRange(i + 1, 18).setValue("DESLIGADO");
    }
  }
  SpreadsheetApp.flush();
}

function sincronizarFeedbacksUniversidade() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetFeedback = ss.getSheetByName("REGISTRO_FEEDBACK");
  var sheetHistorico = ss.getSheetByName("Historico");
  if (!sheetFeedback || !sheetHistorico) return;

  var dadosF = sheetFeedback.getDataRange().getValues();
  if (dadosF.length <= 1) return;

  var mapaFeedbacks = {};
  var estaConcluido = function(val) {
    if (val === undefined || val === null) return false;
    var str = String(val).trim().toUpperCase();
    return str === "1" || str === "OK" || str === "SIM";
  };

  for (var i = 1; i < dadosF.length; i++) {
    var id = String(dadosF[i][0]).trim();
    if (!id) continue;

    var f15Val = dadosF[i][8];
    var f30Val = dadosF[i][9];
    var f45Val = dadosF[i][10];

    if (!mapaFeedbacks[id]) {
      mapaFeedbacks[id] = { f15: false, f30: false, f45: false };
    }

    if (estaConcluido(f15Val)) mapaFeedbacks[id].f15 = true;
    if (estaConcluido(f30Val)) mapaFeedbacks[id].f30 = true;
    if (estaConcluido(f45Val)) mapaFeedbacks[id].f45 = true;
  }

  var rangeH = sheetHistorico.getRange(2, 1, sheetHistorico.getLastRow() - 1, sheetHistorico.getLastColumn());
  var dadosH = rangeH.getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  for (var k = 0; k < dadosH.length; k++) {
    var idHistorico = String(dadosH[k][3]).trim();
    if (!idHistorico) continue;

    var numCols = dadosH[k].length;

    if (mapaFeedbacks[idHistorico]) {
      var fComp = mapaFeedbacks[idHistorico];
      if (fComp.f15) dadosH[k][12] = "OK";
      if (fComp.f30) dadosH[k][13] = "OK";
      if (fComp.f45) dadosH[k][14] = "OK";
    }

    var statusColaborador = (numCols >= 18 && dadosH[k][17]) ? String(dadosH[k][17]).trim().toUpperCase() : "";
    if (statusColaborador === "DESLIGADO") {
      dadosH[k][15] = 0;
      continue;
    }

    var dtAdm = parseDataBR(dadosH[k][6]);
    if (!dtAdm || isNaN(dtAdm.getTime())) continue;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);
    if (diasCasa < 15) {
      dadosH[k][15] = 0;
      continue;
    }

    var f15Val = (numCols >= 13 && dadosH[k][12]) ? String(dadosH[k][12]).trim().toUpperCase() : "";
    var f30Val = (numCols >= 14 && dadosH[k][13]) ? String(dadosH[k][13]).trim().toUpperCase() : "";
    var f45Val = (numCols >= 15 && dadosH[k][14]) ? String(dadosH[k][14]).trim().toUpperCase() : "";

    var atraso = 0;
    if (diasCasa >= 15 && f15Val !== "OK") {
      atraso = diasCasa - 15;
    } else if (diasCasa >= 30 && f30Val !== "OK") {
      atraso = diasCasa - 30;
    } else if (diasCasa >= 45 && f45Val !== "OK") {
      atraso = diasCasa - 45;
    }

    dadosH[k][15] = atraso;
  }

  rangeH.setValues(dadosH);
  SpreadsheetApp.flush();
  console.log("Sincronização e recálculo dos contadores em lote finalizados usando REGISTRO_FEEDBACK.");
}

function enviarRelatorioSemanalCoordenadores() {
  sincronizarFeedbacksUniversidade();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetH = ss.getSheetByName("Historico");
  var sheetF = ss.getSheetByName("Base filiais");
  var dadosH = sheetH.getDataRange().getValues();
  var dadosF = sheetF.getDataRange().getValues();
  var hoje = new Date();
  hoje.setHours(0,0,0,0);

  var coordMap = {};
  for(var j=1; j<dadosF.length; j++) {
    coordMap[String(dadosF[j][0]).replace(/^0+/, '')] = {
      email: dadosF[j][4],
      regional: dadosF[j][2],
      gerenteGP: dadosF[j][5],
      emailRegional: dadosF[j][7]
    };
  }

  var relatorios = {};
  for (var i = 1; i < dadosH.length; i++) {
    var numCols = dadosH[i].length;
    var statusColaborador = (numCols >= 18 && dadosH[i][17]) ? String(dadosH[i][17]).trim().toUpperCase() : "";
    if (statusColaborador === "DESLIGADO") continue;

    var filial = String(dadosH[i][0]).replace(/^0+/, '');
    var info = coordMap[filial];
    if (!info || !info.email) continue;

    var dtAdm = parseDataBR(dadosH[i][6]);
    if (!dtAdm || isNaN(dtAdm.getTime())) continue;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);
    if (diasCasa < 15) continue;

    var f15 = (numCols >= 13 && dadosH[i][12]) ? String(dadosH[i][12]).trim().toUpperCase() : "";
    var f30 = (numCols >= 14 && dadosH[i][13]) ? String(dadosH[i][13]).trim().toUpperCase() : "";
    var f45 = (numCols >= 15 && dadosH[i][14]) ? String(dadosH[i][14]).trim().toUpperCase() : "";

    var temPendenciaReal = false;
    var listaPendencias = [];

    if (diasCasa >= 15 && f15 !== "OK") {
      temPendenciaReal = true;
      listaPendencias.push("Feedback 15 dias");
    }
    if (diasCasa >= 30 && f30 !== "OK") {
      temPendenciaReal = true;
      listaPendencias.push("Feedback 30 dias");
    }
    if (diasCasa >= 45 && f45 !== "OK") {
      temPendenciaReal = true;
      listaPendencias.push("Feedback 45 dias");
    }

    if (temPendenciaReal) {
      var emailDest = info.email.trim();
      if (!relatorios[emailDest]) {
        relatorios[emailDest] = [];
      }
      relatorios[emailDest].push({
        filial: parseInt(filial) || filial,
        regional: info.regional || "Sem Regional",
        nome: dadosH[i][5],
        detalhe: listaPendencias.join(", "),
        gerenteGP: info.gerenteGP,
        emailRegional: info.emailRegional
      });
    }
  }

  for (var email in relatorios) {
    relatorios[email].sort(function(a, b){
      var regA = String(a.regional).toUpperCase();
      var regB = String(b.regional).toUpperCase();
      if (regA < regB) return -1;
      if (regA > regB) return 1;
      return a.filial - b.filial;
    });

    var gruposPorRegional = {};
    relatorios[email].forEach(function(r) {
      var reg = r.regional;
      if (!gruposPorRegional[reg]) {
        gruposPorRegional[reg] = [];
      }
      gruposPorRegional[reg].push(r);
    });

    var totalPendencias = relatorios[email].length;
    var htmlBody = montarEmailRelatorioSemanal(gruposPorRegional, totalPendencias);

    var ccEmails = new Set();
    relatorios[email].forEach(function(item) {
      if (item.gerenteGP) ccEmails.add(String(item.gerenteGP).trim());
      if (item.emailRegional) ccEmails.add(String(item.emailRegional).trim());
    });

    var ccString = Array.from(ccEmails).filter(Boolean).join(",");

    var options = { htmlBody: htmlBody };
    if (ccString) {
      options.cc = ccString;
    }

    try {
      GmailApp.sendEmail(email, "📊 Relatório Semanal de Pendências de Feedback", "", options);
    } catch(e) {
      console.log("Erro ao enviar relatório semanal para " + email + ": " + e.message);
    }
  }
}

// --- BLOCO 3: AUXILIARES E GATILHOS ---

function parseDataBR(s) {
  if (!s) return null;

  var str = "";
  if (s instanceof Date) {
    str = Utilities.formatDate(s, "GMT-3", "dd/MM/yyyy");
  } else {
    str = s.toString().trim();
  }

  var p = str.split("/");
  if (p.length === 3) {
    var dia = parseInt(p[0], 10);
    var mes = parseInt(p[1], 10) - 1;
    var ano = parseInt(p[2], 10);
    if (ano < 100) ano += 2000;
    return new Date(ano, mes, dia);
  }

  var p2 = str.split("-");
  if (p2.length === 3) {
    var ano = parseInt(p2[0], 10);
    if (ano < 100) ano += 2000;
    var mes = parseInt(p2[1], 10) - 1;
    var dia = parseInt(p2[2], 10);
    return new Date(ano, mes, dia);
  }

  return null;
}

function calcularDiasCasa(dataAdmissao, hoje) {
  if (!dataAdmissao) return -1;
  var dAdm = new Date(dataAdmissao.getFullYear(), dataAdmissao.getMonth(), dataAdmissao.getDate());
  var dHoje = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  var diffTime = dHoje.getTime() - dAdm.getTime();
  return Math.round(diffTime / (1000 * 60 * 60 * 24));
}

function normalizarTexto(t) {
  return t ? t.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim() : "";
}

// NOVO: normaliza IDs removendo zeros à esquerda e espaços, para evitar que "123" e "0123"
// sejam tratados como registros diferentes na checagem de duplicidade.
function normalizarID(id) {
  if (id === undefined || id === null) return "";
  var str = String(id).trim();
  var semZeros = str.replace(/^0+/, '');
  return semZeros === "" ? "0" : semZeros;
}

function configurarTodosGatilhos() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) ScriptApp.deleteTrigger(ts[i]);

  // Gatilho Matinal de Admissão configurado às 10:00 AM para garantir que a planilha já tenha sido alimentada
  ScriptApp.newTrigger('rotinaMatinalAgenteIntegrador').timeBased().atHour(10).nearMinute(0).everyDays(1).create();

  // Sincronização automática às 12:00 PM
  ScriptApp.newTrigger('sincronizarFeedbacksUniversidade').timeBased().atHour(12).everyDays(1).create();

  // Cobranças diárias às 13h00 PM
  ScriptApp.newTrigger('rotinaNotificacaoFeedbacks').timeBased().atHour(13).everyDays(1).create();

  // Relatório semanal na segunda-feira às 08h00 AM
  ScriptApp.newTrigger('enviarRelatorioSemanalCoordenadores').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('GP Magalu')
        .addItem('🔄 Sincronizar Feedbacks (REGISTRO_FEEDBACK)', 'sincronizarFeedbacksUniversidade')
        .addToUi();
  } catch (e) {
    console.log("Aviso: Não foi possível carregar o menu.");
  }
}

function onEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();

  if (sheet.getName() !== "Historico") return;

  var col = range.getColumn();
  var row = range.getRow();
  if (row === 1) return;

  if (col === 13 || col === 14 || col === 15) {
    var hoje = new Date();
    hoje.setHours(0,0,0,0);

    var lastCol = sheet.getLastColumn();
    var dadosLinha = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    var statusColaborador = "";
    if (lastCol >= 18 && dadosLinha[17]) {
      statusColaborador = String(dadosLinha[17]).trim().toUpperCase();
    }

    if (statusColaborador === "DESLIGADO") {
      sheet.getRange(row, 16).setValue(0);
      return;
    }

    var dtAdm = parseDataBR(dadosLinha[6]);
    if (!dtAdm || isNaN(dtAdm.getTime())) return;

    var diasCasa = calcularDiasCasa(dtAdm, hoje);

    if (diasCasa < 15) {
      sheet.getRange(row, 16).setValue(0);
      return;
    }

    var f15 = (lastCol >= 13 && dadosLinha[12]) ? String(dadosLinha[12]).trim().toUpperCase() : "";
    var f30 = (lastCol >= 14 && dadosLinha[13]) ? String(dadosLinha[13]).trim().toUpperCase() : "";
    var f45 = (lastCol >= 15 && dadosLinha[14]) ? String(dadosLinha[14]).trim().toUpperCase() : "";

    var atraso = 0;
    if (diasCasa >= 15 && f15 !== "OK") {
      atraso = diasCasa - 15;
    } else if (diasCasa >= 30 && f30 !== "OK") {
      atraso = diasCasa - 30;
    } else if (diasCasa >= 45 && f45 !== "OK") {
      atraso = diasCasa - 45;
    }

    sheet.getRange(row, 16).setValue(atraso);
  }
}

// --- BLOCO 4: TEMPLATES DE E-MAIL (VISUAL) ---
// Todo o HTML de e-mail foi centralizado aqui para ficar mais fácil de manter e ajustar o design
// sem mexer na lógica de negócio das funções acima.

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
  // Monta os badges de etapas (15/30/45 dias) para UM colaborador específico.
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

  // Ordena os colaboradores pelo maior atraso primeiro, para priorizar visualmente os mais urgentes.
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
