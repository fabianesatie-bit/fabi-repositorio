/**
 * SISTEMA GP360 - ATENDIMENTO SOCIAL E CANAL DA MULHER
 * Mapeamento, Governança e Roteamento de Regionais (Apps Script)
 */

const CONFIG = {
  SPREADSHEET_ID: "1InLKT3qmWxAv7N-U1tyoSW0tNI-Qc4LTW0vyza1oSg0",
  SHEET_ROL: "BASE_ROL",
  SHEET_REGISTRO: "BASE_REGISTRO",
  SHEET_HISTORICO: "BASE_HISTORICO",
  SHEET_INTERNOS: "BASE_INTERNOS",
  SHEET_LOJAS: "BASE_LOJAS",
  SHEET_USUARIOS: "BASE_USUSARIOS",
  SHEET_ATENDIMENTO: "BASE_ATENDIMENTO"
};

function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Não foi possível conectar à planilha oficial. Verifique se o ID está correto ou se possui permissão de editor.");
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('GP360 - Atendimento e Apoio')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLoggedUserEmail() {
  return Session.getActiveUser().getEmail();
}

function getUserProfileContext() {
  var email = getLoggedUserEmail().toLowerCase().trim();
  
  if (email === "fabiane.satie@magazineluiza.com.br") {
    return {
      email: "fabiane.satie@magazineluiza.com.br",
      nome: "Fabiane Satie",
      cargo: "Gerente GP (Administradora Geral)",
      diretoria: "Todas",
      regionais: ["TODAS"],
      nivelAcesso: "Administrador"
    };
  }
  
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_USUARIOS);
  if (!sheet) {
    return { email: email, nome: "Usuário Externo", cargo: "Nenhum", diretoria: "Nenhuma", regionais: [], nivelAcesso: "Sem Acesso" };
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().toLowerCase().trim() === email) {
      return {
        email: data[i][0],
        nome: data[i][1],
        cargo: data[i][2],
        diretoria: data[i][3],
        regionais: data[i][4].toString().split(",").map(function(r) { return r.trim().toUpperCase(); }),
        nivelAcesso: data[i][5]
      };
    }
  }
  
  return { email: email, nome: "Usuário Não Cadastrado", cargo: "Nenhum", diretoria: "Nenhuma", regionais: [], nivelAcesso: "Sem Acesso" };
}

function normalizeId(id) {
  if (id === null || id === undefined) return "";
  var cleanStr = id.toString().replace(/[\s\u00A0]+/g, "").trim();
  if (cleanStr.indexOf(".") !== -1) cleanStr = cleanStr.split(".")[0];
  if (cleanStr.indexOf(",") !== -1) cleanStr = cleanStr.split(",")[0];
  return cleanStr.replace(/^0+/, "");
}

function getRegionalByFilial(filialId) {
  var ss = getSpreadsheet();
  var sheetLojas = ss.getSheetByName(CONFIG.SHEET_LOJAS);
  if (!sheetLojas) return "SEM REGIONAL";
  
  var data = sheetLojas.getDataRange().getValues();
  var queryId = normalizeId(filialId);
  
  for (var i = 1; i < data.length; i++) {
    if (normalizeId(data[i][0]) === queryId) {
      return data[i][2].toString().trim().toUpperCase();
    }
  }
  return "SEM REGIONAL";
}

function searchColabById(colabId) {
  var userContext = getUserProfileContext();
  var ss = getSpreadsheet();
  var targetIdNormalized = normalizeId(colabId);
  if (targetIdNormalized === "") return null;

  var colabEncontrado = null;
  var foundInActiveBase = false;
  
  var sheetRol = ss.getSheetByName(CONFIG.SHEET_ROL);
  if (sheetRol) {
    var dataRol = sheetRol.getDataRange().getValues();
    var headers = dataRol[0].map(function(h) { 
      return h.toString().toLowerCase().replace(/[\s\u00A0]+/g, "").trim(); 
    });
    
    var idColIdx = headers.indexOf("id") !== -1 ? headers.indexOf("id") : 1;
    var nomeColIdx = headers.indexOf("nome") !== -1 ? headers.indexOf("nome") : 2;
    var filialColIdx = headers.indexOf("filial") !== -1 ? headers.indexOf("filial") : 3;
    var tempoColIdx = headers.indexOf("tempoempresa") !== -1 ? headers.indexOf("tempoempresa") : 4;
    var cargoColIdx = headers.indexOf("cargo") !== -1 ? headers.indexOf("cargo") : 7;
    
    for (var i = 1; i < dataRol.length; i++) {
      if (normalizeId(dataRol[i][idColIdx]) === targetIdNormalized) {
        var filial = dataRol[i][filialColIdx]; 
        var regional = getRegionalByFilial(filial);
        
        colabEncontrado = {
          id: dataRol[i][idColIdx].toString().trim(), 
          nome: dataRol[i][nomeColIdx].toString().trim(), 
          filial: filial.toString().trim(),
          tempoEmpresa: dataRol[i][tempoColIdx].toString().trim(), 
          regional: regional,
          cargo: dataRol[i][cargoColIdx].toString().trim(), 
          telefone: "",
          situacaoFuncional: "ATIVO"
        };
        foundInActiveBase = true;
        break;
      }
    }
  }
  
  if (!colabEncontrado) {
    var abasFontes = [CONFIG.SHEET_REGISTRO, CONFIG.SHEET_HISTORICO, CONFIG.SHEET_INTERNOS];
    var registrosEncontrados = [];
    
    for (var a = 0; a < abasFontes.length; a++) {
      var sheetVar = ss.getSheetByName(abasFontes[a]);
      if (sheetVar) {
        var dataVar = sheetVar.getDataRange().getValues();
        for (var j = 1; j < dataVar.length; j++) {
          if (normalizeId(dataVar[j][1]) === targetIdNormalized) {
            registrosEncontrados.push({
              id: dataVar[j][1] ? dataVar[j][1].toString().trim() : targetIdNormalized,
              nome: dataVar[j][3] ? dataVar[j][3].toString().trim() : "Nome Indisponível",
              filial: dataVar[j][2] ? dataVar[j][2].toString().trim() : "0000",
              tempoEmpresa: dataVar[j][5] ? dataVar[j][5].toString().trim() : "Desconhecido",
              cargo: dataVar[j][4] ? dataVar[j][4].toString().trim() : "Cargo Anterior",
              telefone: dataVar[j][7] ? dataVar[j][7].toString().trim() : "",
              dataLinha: dataVar[j][0] instanceof Date ? dataVar[j][0] : new Date(0)
            });
          }
        }
      }
    }
    
    if (registrosEncontrados.length > 0) {
      registrosEncontrados.sort(function(x, y) { return y.dataLinha - x.dataLinha; });
      var maisRecente = registrosEncontrados[0];
      
      colabEncontrado = {
        id: maisRecente.id,
        nome: maisRecente.nome,
        filial: maisRecente.filial,
        tempoEmpresa: maisRecente.tempoEmpresa,
        regional: getRegionalByFilial(maisRecente.filial),
        cargo: maisRecente.cargo,
        telefone: maisRecente.telefone,
        situacaoFuncional: "INATIVO"
      };
    }
  }
  
  if (!colabEncontrado) {
    return {
      manualMode: true,
      colaborador: {
        id: targetIdNormalized,
        nome: "",
        filial: "",
        regional: "SÃO PAULO",
        cargo: "",
        tempoEmpresa: "Fora da Base",
        telefone: "",
        situacaoFuncional: "INATIVO"
      },
      historico: []
    };
  }
  
  var regional = colabEncontrado.regional;
  var temPermissao = userContext.nivelAcesso === "Social" || 
                     userContext.nivelAcesso === "Administrador" ||
                     userContext.regionais.indexOf(regional.toUpperCase()) !== -1 ||
                     userContext.regionais.indexOf("TODAS") !== -1;
                     
  if (!temPermissao) {
    return { error: "Acesso Negado: Sua regional (" + userContext.regionais.join(",") + ") não tem permissão para acessar a regional " + regional + "." };
  }
  
  var historico = [];
  var abasVarredura = [
    { nome: CONFIG.SHEET_REGISTRO, fluxo: "SOCIAL/MULHER" },
    { nome: CONFIG.SHEET_HISTORICO, fluxo: "SOCIAL/MULHER" },
    { nome: CONFIG.SHEET_INTERNOS, fluxo: "GP" }
  ];
  
  for (var f = 0; f < abasVarredura.length; f++) {
    var sheet = ss.getSheetByName(abasVarredura[f].nome);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    for (var j = 1; j < data.length; j++) {
      if (normalizeId(data[j][1]) === targetIdNormalized) {
        var dataVal = data[j][0];
        var dataStr = dataVal instanceof Date ? Utilities.formatDate(dataVal, Session.getScriptTimeZone(), "dd/MM/yyyy") : dataVal.toString();
        
        // Identifica subfluxo baseado na observação ou campos se houver (mantendo compatibilidade)
        var fluxoExibicao = abasVarredura[f].fluxo;
        if (fluxoExibicao === "SOCIAL/MULHER") fluxoExibicao = "SOCIAL"; // Default visual fallback
        
        historico.push({
          data: dataStr,
          classificacao: data[j][8].toString().trim().toUpperCase(), 
          obsSocial: data[j][9].toString().trim(),     
          parecerGP: data[j][10].toString().trim(),
          telefone: data[j][7] ? data[j][7].toString().trim() : "",
          dataRetomada: data[j][11] ? data[j][11].toString().trim() : "",
          fluxo: fluxoExibicao
        });
      }
    }
  }
  
  historico.sort(function(a, b) {
    var partesA = a.data.split("/");
    var partesB = b.data.split("/");
    return new Date(partesB[2], partesB[1] - 1, partesB[0]) - new Date(partesA[2], partesA[1] - 1, partesA[0]);
  });
  
  var ultimoTelefone = "";
  for (var k = 0; k < historico.length; k++) {
    if (historico[k].telefone && historico[k].telefone !== "") {
      ultimoTelefone = historico[k].telefone;
      break;
    }
  }
  colabEncontrado.telefone = ultimoTelefone || colabEncontrado.telefone;
  
  return {
    manualMode: !foundInActiveBase,
    colaborador: colabEncontrado,
    historico: historico
  };
}

function getAtendimentosData() {
  var userContext = getUserProfileContext();
  var ss = getSpreadsheet();
  var registrosFiltrados = [];
  
  var ativosMap = {};
  var sheetRol = ss.getSheetByName(CONFIG.SHEET_ROL);
  if (sheetRol) {
    var dataRol = sheetRol.getDataRange().getValues();
    var headersRol = dataRol[0].map(function(h) { return h.toString().toLowerCase().replace(/[\s\u00A0]+/g, "").trim(); });
    var idIdx = headersRol.indexOf("id") !== -1 ? headersRol.indexOf("id") : 1;
    for (var r = 1; r < dataRol.length; r++) {
      var idNormalized = normalizeId(dataRol[r][idIdx]);
      if (idNormalized) {
        ativosMap[idNormalized] = true;
      }
    }
  }

  var sheetLojas = ss.getSheetByName(CONFIG.SHEET_LOJAS);
  var lojaMap = {};
  if (sheetLojas) {
    var dataLojas = sheetLojas.getDataRange().getValues();
    for (var k = 1; k < dataLojas.length; k++) {
      if (dataLojas[k][0]) {
        lojaMap[normalizeId(dataLojas[k][0])] = {
          regional: dataLojas[k][2] ? dataLojas[k][2].toString().trim().toUpperCase() : "SEM REGIONAL",
          diretoria: dataLojas[k][3] ? dataLojas[k][3].toString().trim().toUpperCase() : "SEM DIRETORIA"
        };
      }
    }
  }
  
  var abasFontes = [
    { nome: CONFIG.SHEET_REGISTRO, fluxo: "SOCIAL" },
    { nome: CONFIG.SHEET_HISTORICO, fluxo: "SOCIAL" },
    { nome: CONFIG.SHEET_INTERNOS, fluxo: "GP" }
  ];
  
  for (var a = 0; a < abasFontes.length; a++) {
    var fonte = abasFontes[a];
    var sheet = ss.getSheetByName(fonte.nome);
    if (!sheet) continue;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var idColab = data[i][1] ? data[i][1].toString().trim() : "";
      if (!idColab || idColab === "" || idColab === "null") continue; 
      
      var idNormalized = normalizeId(idColab);
      var filialId = normalizeId(data[i][2]); 
      var infoLoja = lojaMap[filialId] || { regional: "SEM REGIONAL", diretoria: "SEM DIRETORIA" };
      var situacaoFuncional = ativosMap[idNormalized] ? "ATIVO" : "DESLIGADO";
      
      var temPermissao = userContext.nivelAcesso === "Social" || 
                         userContext.nivelAcesso === "Administrador" ||
                         userContext.regionais.indexOf(infoLoja.regional.toUpperCase()) !== -1 ||
                         userContext.regionais.indexOf("TODAS") !== -1;
      
      if (temPermissao) {
        var dataVal = data[i][0];
        var dataStr = dataVal instanceof Date ? Utilities.formatDate(dataVal, Session.getScriptTimeZone(), "dd/MM/yyyy") : dataVal.toString();
        
        registrosFiltrados.push({
          dataEntrada: dataStr,
          id: idColab, 
          colaborador: data[i][3] ? data[i][3].toString().trim() : "",   
          cargo: data[i][4] ? data[i][4].toString().trim() : "",         
          regional: infoLoja.regional,
          diretoria: infoLoja.diretoria,
          coordenador: "Acompanhamento GP360",
          filial: filialId,
          classificacao: data[i][8] ? data[i][8].toString().trim().toUpperCase() : "OUTROS", 
          obsSocial: data[i][9] ? data[i][9].toString().trim() : "",     
          parecerGP: data[i][10] ? data[i][10].toString().trim() : "",
          dataRetomada: data[i][11] ? data[i][11].toString().trim() : "",
          fluxo: fonte.fluxo,
          situacaoFuncional: situacaoFuncional
        });
      }
    }
  }
  
  registrosFiltrados.sort(function(x, y) {
    try {
      var partesX = x.dataEntrada.split("/");
      var partesY = y.dataEntrada.split("/");
      return new Date(partesY[2], partesY[1] - 1, partesY[0]) - new Date(partesX[2], partesX[1] - 1, partesX[0]);
    } catch(e) { return 0; }
  });
  
  return registrosFiltrados;
}

function getClassificacoes() {
  var padrao = ["ACIDENTE TRABALHO", "AFASTAMENTO", "BENEFICIO", "CALAMIDADE", "FINANCEIRO", "PLANO DE SAUDE", "OBITO", "ACOLHIMENTO", "REINTEGRAÇÃO", "SAÚDE", "OUTROS", "PCD", "PROBLEMAS FAMILIARES"];
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_ATENDIMENTO);
    if (!sheet) return padrao;
    
    var data = sheet.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        lista.push(data[i][0].toString().trim().toUpperCase());
      }
    }
    return lista.length > 0 ? lista : padrao;
  } catch (e) {
    return padrao;
  }
}

function addNewClassification(novaClassificacao) {
  var userContext = getUserProfileContext();
  if (userContext.nivelAcesso !== "Administrador") {
    throw new Error("Apenas administradores podem cadastrar novas classificações.");
  }
  
  var novaClassClean = novaClassificacao.toString().trim().toUpperCase();
  if (!novaClassClean) throw new Error("O nome não pode estar em branco.");
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_ATENDIMENTO);
    if (!sheet) throw new Error("Aba BASE_ATENDIMENTO não localizada.");
    sheet.appendRow([novaClassClean]);
  } catch (e) {
    throw new Error("Erro de concorrência ou acesso ao gravar: " + e.message);
  } finally {
    lock.releaseLock();
  }
  return true;
}

function saveNewSocialRecord(payload) {
  var ss = getSpreadsheet();
  var colabInfo = searchColabById(payload.id);
  if (!colabInfo) throw new Error("Colaborador não localizado.");
  
  var colaborador = colabInfo.colaborador;
  var emailLogado = getLoggedUserEmail();
  var situacaoSalvar = payload.situacao || "ATIVO";

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Prevenção contra Race Conditions

    if (payload.fluxo === "GP") {
      var sheetInt = ss.getSheetByName(CONFIG.SHEET_INTERNOS);
      if (!sheetInt) throw new Error("Aba BASE_INTERNOS não localizada.");

      var novaLinhaInt = [
        new Date(),
        colaborador.id,
        colaborador.filial,
        colaborador.nome,
        colaborador.cargo,
        colaborador.tempoEmpresa,
        situacaoSalvar,
        payload.telefone,
        payload.classificacao.toUpperCase(),
        payload.obsSocial,
        payload.parecerGP,
        payload.dataRetomada,
        emailLogado
      ];
      sheetInt.appendRow(novaLinhaInt);

      if (situacaoSalvar === "ATIVO" && payload.dataRetomada) {
        try {
          var dataPartes = payload.dataRetomada.split("-");
          var dataAgenda = new Date(dataPartes[0], dataPartes[1] - 1, dataPartes[2], 9, 0, 0);
          var tituloEvento = "[Retomada GP360] Falar com " + colaborador.nome;
          var descricaoEvento = "Retomar contato com colaborador " + colaborador.nome + " (ID: " + colaborador.id + " / Loja: " + colaborador.filial + ").\n\nClassificação: " + payload.classificacao + "\nObs: " + payload.obsSocial;

          CalendarApp.getDefaultCalendar().createEvent(
            tituloEvento, 
            dataAgenda, 
            new Date(dataAgenda.getTime() + 30 * 60 * 1000), 
            { description: descricaoEvento }
          );
        } catch (e) {
          Logger.log("Erro ao criar evento na agenda: " + e.message);
        }
      }
    } else {
      var sheetReg = ss.getSheetByName(CONFIG.SHEET_REGISTRO);
      if (!sheetReg) throw new Error("Aba BASE_REGISTRO não localizada.");

      var novaLinha = [
        new Date(),
        colaborador.id,
        colaborador.filial,
        colaborador.nome,
        colaborador.cargo,
        colaborador.tempoEmpresa,
        situacaoSalvar,
        payload.telefone,
        payload.classificacao.toUpperCase(),
        payload.obsSocial,
        payload.parecerGP
      ];
      sheetReg.appendRow(novaLinha);
      
      // Roteamento baseado no fluxo escolhido
      if (payload.fluxo === "MULHER") {
        dispararEmailCanalMulher(colaborador, payload);
      } else {
        dispararEmailsPorRegional(colaborador, payload, colabInfo.historico);
      }
    }
  } catch (e) {
    throw new Error("Falha de processamento de gravação: " + e.message);
  } finally {
    lock.releaseLock();
  }
  return true;
}

function dispararEmailsPorRegional(colaborador, payload, historico) {
  var ss = getSpreadsheet();
  var sheetUsuarios = ss.getSheetByName(CONFIG.SHEET_USUARIOS);
  if (!sheetUsuarios) return;
  
  var dataUsr = sheetUsuarios.getDataRange().getValues();
  var destinatarios = [];
  var regionalColab = colaborador.regional.toUpperCase();
  
  for (var i = 1; i < dataUsr.length; i++) {
    var email = dataUsr[i][0] ? dataUsr[i][0].toString().trim() : ""; 
    var cargo = dataUsr[i][2] ? dataUsr[i][2].toString().toUpperCase().trim() : ""; 
    var regionaisDoUsuario = dataUsr[i][4] ? dataUsr[i][4].toString().toUpperCase().split(",").map(function(r) { return r.trim(); }) : []; 
    
    var atendeRegional = regionaisDoUsuario.indexOf(regionalColab) !== -1 || regionaisDoUsuario.indexOf("TODAS") !== -1;
    var deveReceber = (cargo === "SOCIAL") || ((cargo === "COORDENADOR" || cargo === "GERENTEGP" || cargo === "ADMINISTRADOR") && atendeRegional);
    
    if (deveReceber && email && destinatarios.indexOf(email) === -1) {
      destinatarios.push(email);
    }
  }
  
  if (destinatarios.length === 0) return;
  
  var assunto = "[GP360] Notificação Social: Novo Atendimento - " + colaborador.nome + " (" + colaborador.regional + ")";
  var htmlBody = "<div style='font-family: Arial, sans-serif; max-width: 650px; padding: 20px; border: 1px solid #e2e8f0; border-radius: 10px; background-color: #ffffff;'>" +
    "<div style='background-color: #1e1b4b; padding: 15px; border-radius: 8px 8px 0 0; text-align: center;'>" +
      "<h2 style='color: #ffffff; margin: 0; font-size: 18px;'>GP360 Atendimento Social</h2>" +
    "</div>" +
    "<div style='padding: 20px; color: #334155; line-height: 1.6;'>" +
      "<p style='font-size: 14px;'>Um novo atendimento de assistência social foi registrado no sistema.</p>" +
      "<table style='width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px;'>" +
        "<tr><td style='padding: 6px 0; font-weight: bold; width: 35%;'>Colaborador:</td><td style='padding: 6px 0;'>" + colaborador.nome + " (ID: " + colaborador.id + ")</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Cargo / Função:</td><td style='padding: 6px 0;'>" + colaborador.cargo + "</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Regional de Atendimento:</td><td style='padding: 6px 0;'>" + colaborador.regional + " (Filial: " + colaborador.filial + ")</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Contato Telefônico:</td><td style='padding: 6px 0;'>" + payload.telefone + "</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Classificação do Caso:</td><td style='padding: 6px 0;'><span style='background-color: #fef3c7; color: #92400e; padding: 3px 8px; border-radius: 4px; font-weight: bold;'>" + payload.classificacao.toUpperCase() + "</span></td></tr>" +
      "</table>" +
      "<div style='background-color: #f8fafc; padding: 15px; border-radius: 8px; margin-top: 15px; border-left: 4px solid #4f46e5;'>" +
        "<h4 style='margin: 0 0 6px 0; color: #1e293b; font-size: 13px;'>Observações para o Social:</h4>" +
        "<p style='margin: 0; font-size: 12px; color: #475569; white-space: pre-line;'>" + payload.obsSocial + "</p>" +
      "</div>" +
    "</div>" +
  "</div>";
  
  GmailApp.sendEmail(destinatarios.join(","), assunto, "", { htmlBody: htmlBody });
}

function dispararEmailCanalMulher(colaborador, payload) {
  var ss = getSpreadsheet();
  var sheetUsuarios = ss.getSheetByName(CONFIG.SHEET_USUARIOS);
  if (!sheetUsuarios) return;
  
  var dataUsr = sheetUsuarios.getDataRange().getValues();
  var destinatarioTarsila = "";
  
  for (var i = 1; i < dataUsr.length; i++) {
    var email = dataUsr[i][0] ? dataUsr[i][0].toString().trim() : ""; 
    var nome = dataUsr[i][1] ? dataUsr[i][1].toString().toUpperCase().trim() : ""; 
    
    if (nome.indexOf("TARSILA DE PAULA") !== -1 && email) {
      destinatarioTarsila = email;
      break;
    }
  }
  
  if (!destinatarioTarsila) {
    Logger.log("Falha de roteamento: E-mail de Tarsila de Paula não localizado na BASE_USUSARIOS.");
    return;
  }
  
  var assunto = "[GP360 - CANAL DA MULHER] Novo Atendimento Restrito Registrado - " + colaborador.nome;
  var htmlBody = "<div style='font-family: Arial, sans-serif; max-width: 650px; padding: 20px; border: 1px solid #ffe4e6; border-radius: 10px; background-color: #ffffff;'>" +
    "<div style='background-color: #e11d48; padding: 15px; border-radius: 8px 8px 0 0; text-align: center;'>" +
      "<h2 style='color: #ffffff; margin: 0; font-size: 18px;'>🛡️ GP360 - Canal da Mulher</h2>" +
    "</div>" +
    "<div style='padding: 20px; color: #334155; line-height: 1.6;'>" +
      "<p style='font-size: 14px;'>Um novo registro <strong>estrito e confidencial</strong> foi acionado através do Canal da Mulher.</p>" +
      "<table style='width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px;'>" +
        "<tr><td style='padding: 6px 0; font-weight: bold; width: 35%;'>Colaboradora:</td><td style='padding: 6px 0;'>" + colaborador.nome + " (ID: " + colaborador.id + ")</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Cargo / Função:</td><td style='padding: 6px 0;'>" + colaborador.cargo + "</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Regional / Loja:</td><td style='padding: 6px 0;'>" + colaborador.regional + " (Filial: " + colaborador.filial + ")</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Contato Informado:</td><td style='padding: 6px 0;'>" + payload.telefone + "</td></tr>" +
        "<tr><td style='padding: 6px 0; font-weight: bold;'>Classificação:</td><td style='padding: 6px 0;'><span style='background-color: #ffe4e6; color: #9f1239; padding: 3px 8px; border-radius: 4px; font-weight: bold;'>" + payload.classificacao.toUpperCase() + "</span></td></tr>" +
      "</table>" +
      "<div style='background-color: #fff1f2; padding: 15px; border-radius: 8px; margin-top: 15px; border-left: 4px solid #e11d48;'>" +
        "<h4 style='margin: 0 0 6px 0; color: #881337; font-size: 13px;'>Detalhes do Acolhimento:</h4>" +
        "<p style='margin: 0; font-size: 12px; color: #4c0519; white-space: pre-line;'>" + payload.obsSocial + "</p>" +
      "</div>" +
    "</div>" +
  "</div>";
  
  GmailApp.sendEmail(destinatarioTarsila, assunto, "", { htmlBody: htmlBody });
}

function verificarRetomadasDiarias() {
  var ss = getSpreadsheet();
  var sheetInt = ss.getSheetByName(CONFIG.SHEET_INTERNOS);
  if (!sheetInt) return;
  
  var data = sheetInt.getDataRange().getValues();
  var hojeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (var i = 1; i < data.length; i++) {
    var dataRetomadaVal = data[i][11];
    var emailCoordenador = data[i][12];
    var situacao = data[i][6];
    
    if (situacao !== "ATIVO") continue;
    
    var dataRetomadaStr = dataRetomadaVal instanceof Date 
      ? Utilities.formatDate(dataRetomadaVal, Session.getScriptTimeZone(), "yyyy-MM-dd") 
      : (dataRetomadaVal ? dataRetomadaVal.toString().trim() : "");
    
    if (dataRetomadaStr === hojeStr && emailCoordenador) {
      var colabNome = data[i][3];
      var colabId = data[i][1];
      
      MailApp.sendEmail({
        to: emailCoordenador,
        subject: "[GP360] Lembrete de Retomada Social: " + colabNome + " (Hoje)",
        htmlBody: "<p>Olá! Hoje é o dia planejado para retomar contato com o colaborador <strong>" + colabNome + "</strong> (ID: " + colabId + ").</p>"
      });
    }
  }
}
