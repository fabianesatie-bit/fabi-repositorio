/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS - COMPLETO (API + GATILHOS)
 * Atualizado em: 16 de Julho de 2026 (Remetente Personalizado + Filtro Agressivo de Texto)
 */

const CONFIG = {
  nomeAbaBase: 'Extração 1', 
  nomeAbaFiliais: 'Filiais',
  nomeAbaFolgas: 'Folgas Agendadas',
  urlPortal: 'https://script.google.com/a/macros/magazineluiza.com.br/s/AKfycbzWpN0-LZa_Lvd6VugBL5k3tWbHa5jdwKhJyD4UCH4Ml00BeNzved493zspKCgZAmz5Vg/exec',
  emailRemetente: 'gplojas@magazineluiza.com.br' // NOVO: E-mail de remetente desejado
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Banco de Horas')
    .addItem('Configurar Gatilhos (Automático)', 'configurarGatilhos')
    .addToUi();
}

/**
 * FUNÇÃO OBRIGATÓRIA DO WEB APP (PORTAL)
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal - Banco de Horas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 1. CONFIGURAÇÃO DE GATILHOS AUTOMÁTICOS
 */
function configurarGatilhos() {
  const ts = ScriptApp.getProjectTriggers();
  ts.forEach(t => ScriptApp.deleteTrigger(t));
  
  // Segunda-feira 08h - Relatório Crítico (Regionais/Coords)
  ScriptApp.newTrigger('enviarRelatorioSegunda').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  
  // Terça e Quinta 08h - Lembrete para o Coordenador disparar e-mails
  ScriptApp.newTrigger('lembreteCoordenadorTercaQuinta').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(8).create();
  ScriptApp.newTrigger('lembreteCoordenadorTercaQuinta').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(8).create();
  
  // Diário 08h - Lembrete de folga no dia seguinte
  ScriptApp.newTrigger('enviarLembreteFolgaAmanha').timeBased().everyDays(1).atHour(8).create();

  verificarAbaFolgas();

  // Proteção contra erro de contexto ao executar no editor
  try {
    SpreadsheetApp.getUi().alert('✅ Gatilhos configurados! Sistema a operar no modo automático.');
  } catch (e) {
    Logger.log('✅ Gatilhos configurados com sucesso!');
  }
}

function verificarAbaFolgas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG.nomeAbaFolgas);
  if (!aba) {
    aba = ss.insertSheet(CONFIG.nomeAbaFolgas);
    aba.appendRow(["Carimbo de Data/Hora", "Filial", "Nome", "ID do Colaborador", "Data da Folga"]);
    aba.getRange("A1:E1").setFontWeight("bold").setBackground("#d9ead3");
  }
}

function parseHoras(valor) {
  if (valor === "" || valor === null || valor === undefined) return 0;
  if (typeof valor === 'number') return valor;
  
  let strVal = valor.toString().trim();
  if (strVal.includes(',')) {
    strVal = strVal.replace(/\./g, ''); 
    strVal = strVal.replace(',', '.');  
  }
  return parseFloat(strVal) || 0;
}

// NOVO: Função de envio de e-mail centralizada (Lida com o remetente gplojas)
function enviarEmail(para, assunto, corpoHtml, cc = "") {
  let opcoes = { htmlBody: corpoHtml };
  if (cc) opcoes.cc = cc;

  try {
    // Pega os aliases (emails configurados no "Enviar como") da conta atual
    const aliases = GmailApp.getAliases();
    
    // Se a conta tem permissão para mandar como gplojas, usa ele.
    if (aliases.includes(CONFIG.emailRemetente)) {
      opcoes.from = CONFIG.emailRemetente;
    }
    
    GmailApp.sendEmail(para, assunto, "", opcoes);
  } catch (e) {
    Logger.log("Erro ao enviar email para " + para + ": " + e.message);
  }
}

// ATUALIZADO: Filtro Agressivo (Remove espaços, hífens e acentos para garantir a comparação)
function padronizarTexto(texto) {
  if (!texto) return "";
  return texto.toString()
    .trim() 
    .toUpperCase() 
    .normalize("NFD") 
    .replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/[^A-Z0-9]/g, ""); // Remove TUDO que não for letra ou número (espaços, hífens, etc)
}

/**
 * ============================================================================
 * COMUNICAÇÃO COM O PORTAL WEB (FRONTEND)
 * ============================================================================
 */

function getRegionaisPortal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.nomeAbaFiliais);
    if (!aba) return ["⚠️ ERRO: Aba '" + CONFIG.nomeAbaFiliais + "' não existe!"];
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return ["⚠️ ERRO: Aba '" + CONFIG.nomeAbaFiliais + "' está vazia!"];
    
    // Devolve os nomes originais (com hífen e espaço) para o menu dropdown do portal
    let regionais = [...new Set(dados.slice(1).map(r => r[2]).filter(String))];
    if (regionais.length === 0) return ["⚠️ ERRO: Nenhuma regional na coluna C"];
    
    return regionais.sort();
  } catch (e) {
    return ["⚠️ ERRO FATAL: " + e.message];
  }
}

function getColaboradoresPortal(tipo, parametro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.nomeAbaBase);
    const abaFolgas = ss.getSheetByName(CONFIG.nomeAbaFolgas);
    
    if (!aba) throw new Error("Aba '" + CONFIG.nomeAbaBase + "' não encontrada.");
    
    let folgasMap = {};
    if (abaFolgas) {
      const dadosFolgas = abaFolgas.getDataRange().getValues();
      for (let i = 1; i < dadosFolgas.length; i++) {
        let idColab = dadosFolgas[i][3].toString().trim();
        let dataFolga = dadosFolgas[i][4];
        
        if (dataFolga instanceof Date) {
          folgasMap[idColab] = Utilities.formatDate(dataFolga, "GMT-3", "dd/MM/yyyy");
        } else {
          let strData = dataFolga.toString().trim();
          let parts = strData.split('-');
          if (parts.length === 3) {
            folgasMap[idColab] = `${parts[2]}/${parts[1]}/${parts[0]}`;
          } else {
            folgasMap[idColab] = strData;
          }
        }
      }
    }
    
    const dadosBase = aba.getDataRange().getValues();
    let resultados = [];
    
    // Limpa a string da regional que veio do Portal (ex: "SP-ABC LITORAL" vira "SPABCLITORAL")
    let paramLimpo = padronizarTexto(parametro);

    for (let i = 1; i < dadosBase.length; i++) {
      let row = dadosBase[i];
      let filial = row[0].toString().trim().padStart(4, '0');
      
      // Limpa a string da regional que está na base
      let regional = padronizarTexto(row[3]); 
      let id = row[4];
      let nome = row[5];
      let horas = parseHoras(row[8]); 
      
      if (horas >= 5) {
        if ((tipo === 'coord' && regional === paramLimpo) || (tipo === 'gerente' && filial === paramLimpo)) {
          resultados.push({ 
            filial: filial, 
            regional: row[3] ? row[3].toString().trim() : "", 
            id: id.toString().trim(), 
            nome: nome, 
            horas: horas,
            dataAgendada: folgasMap[id.toString().trim()] || null
          });
        }
      }
    }
    return resultados;
  } catch (e) {
    return [{ filial: 'ERRO', regional: 'ERRO', id: '000', nome: 'Falha: ' + e.message, horas: 0, dataAgendada: null }];
  }
}

function dispararParaGerentes(regionalSelecionada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dadosBase = ss.getSheetByName(CONFIG.nomeAbaBase).getDataRange().getValues();
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  const appUrl = CONFIG.urlPortal; 
  
  let agrupamentoLoja = {};
  let regSelecionadaLimpa = padronizarTexto(regionalSelecionada);
  
  for (let i = 1; i < dadosBase.length; i++) {
    let filial = dadosBase[i][0].toString().trim().padStart(4, '0');
    let regional = padronizarTexto(dadosBase[i][3]); 
    let nome = dadosBase[i][5];
    let horas = parseHoras(dadosBase[i][8]);
    
    if (regional === regSelecionadaLimpa && horas >= 5) {
      if (!agrupamentoLoja[filial]) agrupamentoLoja[filial] = [];
      agrupamentoLoja[filial].push({ nome: nome, horas: horas });
    }
  }

  let lojasProcessadas = 0;
  Object.keys(agrupamentoLoja).forEach(filial => {
    let filialInfo = dadosFiliais.find(f => f[0].toString().trim().padStart(4, '0') === filial);
    if (filialInfo) {
      let emailLoja = filialInfo[3];
      let magicLink = `${appUrl}?view=gerente&filial=${filial}`; 
      
      let tabelaHtml = agrupamentoLoja[filial].map(c => {
        let tag = c.horas > 10 ? `<span style="color:red; font-weight:bold;">CRÍTICO - URGENTE</span>` : "Combinar Folga";
        return `<tr><td style="padding: 8px; border-bottom: 1px solid #ddd;">${c.nome}</td><td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${c.horas}</td><td style="padding: 8px; border-bottom: 1px solid #ddd; text-align: center;">${tag}</td></tr>`;
      }).join('');
      
      let corpoHtml = `
        <div style="font-family: sans-serif; max-width: 600px;">
          <h2 style="color: #e67e22;">Ação Requerida: Banco de Horas - Filial ${filial}</h2>
          <p>Você possui colaboradores a exceder o limite de banco de horas permitido.</p>
          <table style="border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 14px;">
            <tr style="background-color: #f2f2f2;"><th style="padding: 10px; text-align: left;">Colaborador</th><th>Horas</th><th>Orientação</th></tr>
            ${tabelaHtml}
          </table>
          <div style="text-align: center; margin-top: 30px;">
            <a href="${magicLink}" style="background-color: #27ae60; color: white; padding: 14px 28px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 16px;">AGENDAR FOLGAS NO PORTAL</a>
          </div>
        </div>`;
      
      enviarEmail(emailLoja, `Orientação Banco de Horas - Filial ${filial}`, corpoHtml);
      lojasProcessadas++;
    }
  });
  
  return { success: true, count: lojasProcessadas };
}

function salvarFolgasPortal(filial, agendamentos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(CONFIG.nomeAbaFolgas);
  verificarAbaFolgas();
  
  let agora = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
  
  agendamentos.forEach(ag => {
    aba.appendRow([agora, filial, ag.nome, ag.id, ag.dataFolga]);
  });
  return true;
}

/**
 * ============================================================================
 * ROTINAS AUTOMÁTICAS DE E-MAIL
 * ============================================================================
 */

function lembreteCoordenadorTercaQuinta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  const appUrl = CONFIG.urlPortal;
  let emailsCoordenadores = [...new Set(dadosFiliais.slice(1).map(row => row[4]).filter(String))];
  
  let corpoHtml = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; padding: 20px; border: 1px solid #ddd; border-radius: 8px;">
      <h2 style="color: #2c3e50;">Bom dia! Ação Necessária ⚠️</h2>
      <p>Hoje é dia de notificar os gerentes da sua regional sobre o Banco de Horas pendente.</p>
      <div style="text-align: center; margin: 30px 0;">
        <a href="${appUrl}" style="background-color: #0056b3; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;">Aceder Portal Banco de Horas</a>
      </div>
    </div>`;
  
  emailsCoordenadores.forEach(email => enviarEmail(email, "Lembrete: Disparo de Notificações - Banco de Horas", corpoHtml));
}

function enviarRelatorioSegunda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dadosBase = ss.getSheetByName(CONFIG.nomeAbaBase).getDataRange().getValues();
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  
  let relatorio = {}; 
  for (let i = 1; i < dadosBase.length; i++) {
    let filial = dadosBase[i][0].toString().trim().padStart(4, '0');
    let regional = dadosBase[i][3];
    let horas = parseHoras(dadosBase[i][8]);
    
    if (horas >= 5) {
      let filialInfo = dadosFiliais.find(f => f[0].toString().trim().padStart(4, '0') === filial);
      if (filialInfo) {
        let emailCoord = filialInfo[4];
        let emailReg = filialInfo[7];
        let chaveDestino = emailCoord + "," + emailReg; 
        
        if (!relatorio[chaveDestino]) relatorio[chaveDestino] = {};
        if (!relatorio[chaveDestino][regional]) relatorio[chaveDestino][regional] = [];
        
        relatorio[chaveDestino][regional].push({ filial: filial, nome: dadosBase[i][5], horas: horas, critico: horas > 10 });
      }
    }
  }

  Object.keys(relatorio).forEach(destinatarios => {
    let corpoHtml = `<div style="font-family: Arial, sans-serif;"><h2>Resumo Semanal: Banco de Horas</h2>`;
    Object.keys(relatorio[destinatarios]).forEach(regional => {
      let lista = relatorio[destinatarios][regional];
      corpoHtml += `<h3 style="color: #0056b3;">Regional: ${regional}</h3><table border="1" style="border-collapse: collapse; width: 100%;">`;
      corpoHtml += `<tr style="background-color: #f4f4f4;"><th>Filial</th><th>Colaborador</th><th>Horas</th><th>Risco</th></tr>`;
      lista.forEach(p => {
        let cor = p.critico ? "background-color: #ffeaea; color: #d9534f;" : "";
        corpoHtml += `<tr style="${cor}"><td align="center">${p.filial}</td><td>${p.nome}</td><td align="center">${p.horas}</td><td align="center">${p.critico ? "CRÍTICO" : "Atenção"}</td></tr>`;
      });
      corpoHtml += `</table><br>`;
    });
    corpoHtml += `</div>`;
    
    let emails = destinatarios.split(",");
    enviarEmail(emails[0], "Relatório Crítico - Banco de Horas", corpoHtml, emails[1]);
  });
}

function enviarLembreteFolgaAmanha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaFolgas = ss.getSheetByName(CONFIG.nomeAbaFolgas);
  if (!abaFolgas) return;
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  const dadosFolgas = abaFolgas.getDataRange().getValues();
  
  const amanha = new Date();
  amanha.setDate(amanha.getDate() + 1);
  const dataAlvo = Utilities.formatDate(amanha, "GMT-3", "yyyy-MM-dd"); 
  
  for (let i = 1; i < dadosFolgas.length; i++) {
    let dataFolgaRegistrada = dadosFolgas[i][4]; 
    if (dataFolgaRegistrada === dataAlvo) {
      let filial = dadosFolgas[i][1].toString().padStart(4, '0');
      let fInfo = dadosFiliais.find(f => f[0].toString().padStart(4, '0') === filial);
      if (fInfo) {
        let emailLoja = fInfo[3];
        let nome = dadosFolgas[i][2];
        enviarEmail(emailLoja, `Lembrete: Folga Amanhã - Filial ${filial}`, `Olá! Lembrete de que o colaborador ${nome} possui folga agendada para amanhã, conforme registado no portal.`);
      }
    }
  }
}
