/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS - COMPLETO (API + GATILHOS)
 * Atualizado em: 17 de Setembro de 2026
 * Correção: Suporte a filtro por Regional, Diretoria e Filial + Mapeamento da Extração 1
 */

const CONFIG = {
  nomeAbaBase: 'Extração 1', 
  nomeAbaFiliais: 'Filiais',
  nomeAbaFolgas: 'Folgas Agendadas',
  nomeAbaDadosAtivos: 'DADOS_ATIVOS', // Aba de situação/status dos colaboradores
  urlPortal: 'https://script.google.com/a/macros/magazineluiza.com.br/s/AKfycbzWpN0-LZa_Lvd6VugBL5k3tWbHa5jdwKhJyD4UCH4Ml00BeNzved493zspKCgZAmz5Vg/exec',
  emailRemetente: 'gplojas@magazineluiza.com.br'
};

/**
 * Normaliza o ID da filial aplicando a regra > 3000 e formatação com 4 dígitos
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  let num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num).padStart(4, '0');
}

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
  
  // Segunda-feira 08h - Relatório Crítico
  ScriptApp.newTrigger('enviarRelatorioSegunda').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  
  // Terça e Quinta 08h - Lembrete Coordenador
  ScriptApp.newTrigger('lembreteCoordenadorTercaQuinta').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(8).create();
  ScriptApp.newTrigger('lembreteCoordenadorTercaQuinta').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.THURSDAY).atHour(8).create();
  
  // Diário 08h - Lembrete de Folga Amanhã
  ScriptApp.newTrigger('enviarLembreteFolgaAmanha').timeBased().everyDays(1).atHour(8).create();

  verificarAbaFolgas();

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

/**
 * Leitura e conversão de Horas
 */
function parseHoras(valor) {
  if (valor === "" || valor === null || valor === undefined) return 0;
  
  if (valor instanceof Date) {
    return valor.getHours() + (valor.getMinutes() / 60);
  }
  
  if (typeof valor === 'number') return valor;
  
  let strVal = valor.toString().trim();
  
  if (strVal.includes(':')) {
    let partes = strVal.split(':');
    let h = parseFloat(partes[0]) || 0;
    let m = parseFloat(partes) || 0;
    return h + (m / 60);
  }
  
  if (strVal.includes(',')) {
    strVal = strVal.replace(/\./g, '').replace(',', '.');
  }
  return parseFloat(strVal) || 0;
}

// Envio de e-mail centralizado
function enviarEmail(para, assunto, corpoHtml, cc = "") {
  let opcoes = { htmlBody: corpoHtml };
  if (cc) opcoes.cc = cc;

  try {
    const aliases = GmailApp.getAliases();
    if (aliases.includes(CONFIG.emailRemetente)) {
      opcoes.from = CONFIG.emailRemetente;
    }
    
    GmailApp.sendEmail(para, assunto, "", opcoes);
  } catch (e) {
    Logger.log("Erro ao enviar email para " + para + ": " + e.message);
  }
}

/**
 * Padronização de textos para comparação segura
 */
function padronizarTexto(texto) {
  if (!texto) return "";
  let s = texto.toString().trim().toUpperCase();
  s = s.replace(/REGIONAL\s*\d*\s*[-–]?\s*/g, '');
  s = s.replace(/DIRETORIA\s*\d*\s*[-–]?\s*/g, '');
  return s
    .normalize("NFD") 
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

/**
 * Mapeia o ID do colaborador para a sua Situação a partir da aba DADOS_ATIVOS.
 */
function carregarStatusColaboradores(ss) {
  let statusMap = {};
  const abaDadosAtivos = ss.getSheetByName(CONFIG.nomeAbaDadosAtivos);
  if (abaDadosAtivos) {
    const dadosAtivos = abaDadosAtivos.getDataRange().getValues();
    for (let i = 1; i < dadosAtivos.length; i++) {
      let row = dadosAtivos[i];
      let id = row !== undefined && row !== null ? row.toString().trim() : "";
      let situacao = row[6] !== undefined && row[6] !== null ? row[6].toString().trim() : "";
      if (id) {
        statusMap[id] = situacao;
      }
    }
  }
  return statusMap;
}

/**
 * Validação de colaborador ativo
 */
function isEmAtividadeNormal(situacao) {
  if (!situacao) return true;
  let s = situacao.toString().trim().toUpperCase();
  return s === "" || s === "EM ATIVIDADE NORMAL" || s === "ATIVO" || s.includes("ATIVIDADE NORMAL");
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
    
    // Lista todas as opções da Coluna B (Regionais) e Coluna C (Diretorias se houver)
    let regionais = [...new Set(dados.slice(1).map(r => r).filter(String))];
    let diretorias = [...new Set(dados.slice(1).map(r => r[2]).filter(String))];
    
    let listaFinal = [...new Set([...regionais, ...diretorias])];
    
    if (listaFinal.length === 0) return ["⚠️ ERRO: Nenhuma regional encontrada"];
    
    return listaFinal.sort();
  } catch (e) {
    return ["⚠️ ERRO FATAL: " + e.message];
  }
}

function getColaboradoresPortal(tipo, parametro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.nomeAbaBase);
    const abaFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais);
    const abaFolgas = ss.getSheetByName(CONFIG.nomeAbaFolgas);
    
    if (!aba) throw new Error("Aba '" + CONFIG.nomeAbaBase + "' não encontrada.");
    
    const statusMap = carregarStatusColaboradores(ss);
    
    // Mapeia filiais para suas respectivas Regionais e Diretorias
    let mapaFilialDiretoria = {};
    let mapaFilialRegional = {};
    if (abaFiliais) {
      let dadosF = abaFiliais.getDataRange().getValues();
      for (let f = 1; f < dadosF.length; f++) {
        let fId = normalizarFilialId(dadosF[f][0]);
        mapaFilialRegional[fId] = padronizarTexto(dadosF[f]); // Coluna B
        mapaFilialDiretoria[fId] = padronizarTexto(dadosF[f][2]); // Coluna C
      }
    }
    
    let folgasMap = {};
    if (abaFolgas) {
      const dadosFolgas = abaFolgas.getDataRange().getValues();
      for (let i = 1; i < dadosFolgas.length; i++) {
        let idColab = dadosFolgas[i][3] ? dadosFolgas[i][3].toString().trim() : "";
        let dataFolga = dadosFolgas[i][4];
        
        if (dataFolga instanceof Date) {
          folgasMap[idColab] = Utilities.formatDate(dataFolga, "GMT-3", "dd/MM/yyyy");
        } else if (dataFolga) {
          let strData = dataFolga.toString().trim();
          let parts = strData.split('-');
          if (parts.length === 3) {
            folgasMap[idColab] = `${parts[2]}/${parts}/${parts[0]}`;
          } else {
            folgasMap[idColab] = strData;
          }
        }
      }
    }
    
    const dadosBase = aba.getDataRange().getValues();
    let resultados = [];
    
    let paramLimpo = padronizarTexto(parametro);
    let filialParamNorm = normalizarFilialId(parametro);

    for (let i = 1; i < dadosBase.length; i++) {
      let row = dadosBase[i];
      
      let filial        = normalizarFilialId(row[0]);                           // Coluna A
      let id            = row ? row.toString().trim() : "";               // Coluna B
      let nome          = row[2] ? row[2].toString().trim() : "";               // Coluna C
      let regionalOrig  = row[3] ? row[3].toString().trim() : "";               // Coluna D
      let regional      = padronizarTexto(regionalOrig);                        // Coluna D limpa
      let situacaoLinha = row[6] ? row[6].toString().trim() : "";               // Coluna G
      let horas         = parseHoras(row[7]);                                   // Coluna H
      
      let situacaoFinal = statusMap[id] || situacaoLinha;

      if (!isEmAtividadeNormal(situacaoFinal)) {
        continue;
      }

      if (horas >= 5) {
        let match = false;
        
        if (tipo === 'gerente' && filial === filialParamNorm) {
          match = true;
        } else if (tipo === 'coord') {
          let regFilial = mapaFilialRegional[filial] || "";
          let dirFilial = mapaFilialDiretoria[filial] || "";
          
          // Confere se bate com Regional da Extração 1, Regional da aba Filiais ou Diretoria da aba Filiais
          if (regional === paramLimpo || regional.includes(paramLimpo) || paramLimpo.includes(regional) ||
              regFilial === paramLimpo || regFilial.includes(paramLimpo) || paramLimpo.includes(regFilial) ||
              dirFilial === paramLimpo || dirFilial.includes(paramLimpo) || paramLimpo.includes(dirFilial)) {
            match = true;
          }
        }

        if (match) {
          resultados.push({ 
            filial: filial, 
            regional: regionalOrig, 
            id: id, 
            nome: nome, 
            horas: horas,
            dataAgendada: folgasMap[id] || null
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
  
  const statusMap = carregarStatusColaboradores(ss);
  
  let mapaFilialDiretoria = {};
  let mapaFilialRegional = {};
  if (dadosFiliais) {
    for (let f = 1; f < dadosFiliais.length; f++) {
      let fId = normalizarFilialId(dadosFiliais[f][0]);
      mapaFilialRegional[fId] = padronizarTexto(dadosFiliais[f]);
      mapaFilialDiretoria[fId] = padronizarTexto(dadosFiliais[f][2]);
    }
  }

  let agrupamentoLoja = {};
  let regSelecionadaLimpa = padronizarTexto(regionalSelecionada);
  
  for (let i = 1; i < dadosBase.length; i++) {
    let row = dadosBase[i];
    let filial        = normalizarFilialId(row[0]);
    let id            = row ? row.toString().trim() : "";
    let nome          = row[2] ? row[2].toString().trim() : "";
    let regional      = padronizarTexto(row[3]);
    let situacaoLinha = row[6] ? row[6].toString().trim() : "";
    let horas         = parseHoras(row[7]);
    
    let situacaoFinal = statusMap[id] || situacaoLinha;

    if (!isEmAtividadeNormal(situacaoFinal)) {
      continue;
    }
    
    let regFilial = mapaFilialRegional[filial] || "";
    let dirFilial = mapaFilialDiretoria[filial] || "";

    let match = (regional === regSelecionadaLimpa || regional.includes(regSelecionadaLimpa) || regSelecionadaLimpa.includes(regional) ||
                 regFilial === regSelecionadaLimpa || regFilial.includes(regSelecionadaLimpa) || regSelecionadaLimpa.includes(regFilial) ||
                 dirFilial === regSelecionadaLimpa || dirFilial.includes(regSelecionadaLimpa) || regSelecionadaLimpa.includes(dirFilial));
    
    if (match && horas >= 5) {
      if (!agrupamentoLoja[filial]) agrupamentoLoja[filial] = [];
      agrupamentoLoja[filial].push({ nome: nome, horas: horas });
    }
  }

  let lojasProcessadas = 0;
  Object.keys(agrupamentoLoja).forEach(filial => {
    let filialInfo = dadosFiliais.find(f => normalizarFilialId(f[0]) === filial);
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
  let filialNorm = normalizarFilialId(filial);

  agendamentos.forEach(ag => {
    aba.appendRow([agora, filialNorm, ag.nome, ag.id, ag.dataFolga]);
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
        <a href="${appUrl}" style="background-color: #0056b3; color: white; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;">Acessar Portal Banco de Horas</a>
      </div>
    </div>`;
  
  emailsCoordenadores.forEach(email => enviarEmail(email, "Lembrete: Disparo de Notificações - Banco de Horas", corpoHtml));
}

function enviarRelatorioSegunda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dadosBase = ss.getSheetByName(CONFIG.nomeAbaBase).getDataRange().getValues();
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  
  const statusMap = carregarStatusColaboradores(ss);
  
  let relatorio = {}; 
  for (let i = 1; i < dadosBase.length; i++) {
    let row = dadosBase[i];
    let filial        = normalizarFilialId(row[0]);
    let id            = row ? row.toString().trim() : "";
    let nome          = row[2] ? row[2].toString().trim() : "";
    let regional      = row[3];
    let situacaoLinha = row[6] ? row[6].toString().trim() : "";
    let horas         = parseHoras(row[7]);
    
    let situacaoFinal = statusMap[id] || situacaoLinha;

    if (!isEmAtividadeNormal(situacaoFinal)) {
      continue;
    }
    
    if (horas >= 5) {
      let filialInfo = dadosFiliais.find(f => normalizarFilialId(f[0]) === filial);
      if (filialInfo) {
        let emailCoord = filialInfo[4];
        let emailReg = filialInfo[7];
        let chaveDestino = emailCoord + (emailReg ? ("," + emailReg) : ""); 
        
        if (!relatorio[chaveDestino]) relatorio[chaveDestino] = {};
        if (!relatorio[chaveDestino][regional]) relatorio[chaveDestino][regional] = [];
        
        relatorio[chaveDestino][regional].push({ filial: filial, nome: nome, horas: horas, critico: horas > 10 });
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
    enviarEmail(emails[0], "Relatório Crítico - Banco de Horas", corpoHtml, emails || "");
  });
}

function enviarLembreteFolgaAmanha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaFolgas = ss.getSheetByName(CONFIG.nomeAbaFolgas);
  if (!abaFolgas) return;
  const dadosFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais).getDataRange().getValues();
  const dadosFolgas = abaFolgas.getDataRange().getValues();
  
  const statusMap = carregarStatusColaboradores(ss);
  
  const amanha = new Date();
  amanha.setDate(amanha.getDate() + 1);
  const dataAlvo = Utilities.formatDate(amanha, "GMT-3", "yyyy-MM-dd"); 
  
  for (let i = 1; i < dadosFolgas.length; i++) {
    let dataFolgaRegistrada = dadosFolgas[i][4]; 
    if (dataFolgaRegistrada === dataAlvo) {
      let idColab = dadosFolgas[i][3] ? dadosFolgas[i][3].toString().trim() : "";
      
      let situacaoFinal = statusMap[idColab];
      if (situacaoFinal !== undefined && !isEmAtividadeNormal(situacaoFinal)) {
        continue;
      }

      let filial = normalizarFilialId(dadosFolgas[i]);
      let fInfo = dadosFiliais.find(f => normalizarFilialId(f[0]) === filial);
      if (fInfo) {
        let emailLoja = fInfo[3];
        let nome = dadosFolgas[i][2];
        enviarEmail(emailLoja, `Lembrete: Folga Amanhã - Filial ${filial}`, `Olá! Lembrete de que o colaborador ${nome} possui folga agendada para amanhã, conforme registrado no portal.`);
      }
    }
  }
}
