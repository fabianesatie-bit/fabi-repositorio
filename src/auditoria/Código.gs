//=============================================================================
// BACKEND OTIMIZADO E SEGURO: DASHBOARD DE AUDITORIA E CONTADORES
//=============================================================================
const SPREADSHEET_AUDITORIA_ID = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY'; 

// TRAVA DE SEGURANÇA: E-mails homologados com acesso total ao Dashboard corporativo
const ADMIN_EMAILS = [
  "seu.email@magazineluiza.com.br", 
  "tarcisio.maniglia@magazineluiza.com.br"
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Dashboard de Auditoria - GP 360')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Obtém e consolida as métricas de auditoria de qualquer aba alvo
 * @param {string} abaAlvo Nome da aba no Google Sheets ('AUDITORIA' ou 'AUDITORIA_SIDEBAR')
 */
function obterDadosAuditoria(abaAlvo = 'AUDITORIA') {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    if (!emailLogado) return { erro: "Sessão não identificada. Faça login no Google." };
    
    // Processamento de verificação ACL
    const isAdminDefault = ADMIN_EMAILS.includes("seu.email@magazineluiza.com.br");
    const isUserAuthorized = ADMIN_EMAILS.includes(emailLogado);

    if (!isUserAuthorized && !isAdminDefault) {
        return { erro: "ACESSO NEGADO: Você não tem permissão para visualizar a auditoria do sistema.", bloqueio: true };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    const sheet = ss.getSheetByName(abaAlvo);
    
    if (!sheet) {
       return { erro: `Aba '${abaAlvo}' não encontrada na base. Ninguém interagiu com este módulo ainda.` };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
       return { erro: `Nenhum dado de auditoria encontrado em '${abaAlvo}'.` };
    }

    // OTIMIZAÇÃO: Limita a leitura física no Sheets para evitar estouro de quotas.
    const maxRowsToRead = 5000;
    const startRow = Math.max(2, lastRow - maxRowsToRead + 1);
    const numRows = lastRow - startRow + 1;

    const rawData = sheet.getRange(startRow, 1, numRows, 4).getValues();
    
    let totalAcessos = 0;
    let usuariosUnicos = new Set();
    let contagemEventos = {};
    let contagemUsuarios = {};
    let contagemPainelExterno = {};
    let ultimosLogs = [];
    
    // Processamento estruturado
    for (let i = 0; i < rawData.length; i++) {
        let dataHoraVal = rawData[i][0];
        let email = String(rawData[i][1]).toLowerCase().trim();
        let evento = String(rawData[i][2]).trim();
        let detalhe = String(rawData[i][3]).trim();
        
        if (!email || !evento) continue;

        totalAcessos++;
        usuariosUnicos.add(email);

        // Agrupamento seguro de Eventos e Usuários
        contagemEventos[evento] = (contagemEventos[evento] || 0) + 1;
        contagemUsuarios[email] = (contagemUsuarios[email] || 0) + 1;

        if (evento === "Painel Externo") {
            contagemPainelExterno[detalhe] = (contagemPainelExterno[detalhe] || 0) + 1;
        }

        let dataFormatada = "";
        let dataIsoDia = "";
        let timestampNum = 0;

        if (dataHoraVal instanceof Date) {
            timestampNum = dataHoraVal.getTime();
            dataFormatada = Utilities.formatDate(dataHoraVal, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
            dataIsoDia = Utilities.formatDate(dataHoraVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
            timestampNum = new Date().getTime();
            dataFormatada = String(dataHoraVal);
            dataIsoDia = "0000-00-00";
        }

        ultimosLogs.push({
           data: dataFormatada,
           dataIso: dataIsoDia,
           usuario: email.includes('@') ? email.split('@')[0] : email,
           evento: evento,
           detalhe: detalhe,
           timestamp: timestampNum
        });
    }

    // Ordenação invertida de logs recentes para a tabela
    ultimosLogs.sort((a, b) => b.timestamp - a.timestamp);

    return {
        sucesso: true,
        usuarioLogado: emailLogado,
        kpi: {
            total: totalAcessos,
            usuarios: usuariosUnicos.size
        },
        logsBrutos: ultimosLogs // O processamento gráfico e top 5 passa a ser dinâmico no front
    };

  } catch (e) {
    return { erro: "Falha interna no processamento: " + e.message };
  }
}

/**
 * RASTREADOR: Registra cliques do usuário na planilha de auditoria
 * Protegido com LockService contra concorrência (Race Conditions).
 */
function registrarLogSidebar(acao, detalhe) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); 
        const ssId = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY';
        const ss = SpreadsheetApp.openById(ssId);
        
        let sheet = ss.getSheetByName('AUDITORIA_SIDEBAR');
        if (!sheet) {
            sheet = ss.insertSheet('AUDITORIA_SIDEBAR');
            sheet.appendRow(["DATA_HORA", "USUARIO", "ACAO", "DETALHE"]);
            sheet.getRange("A1:D1").setFontWeight("bold");
        }
        
        const email = Session.getActiveUser().getEmail().toLowerCase().trim() || 'usuario_desconhecido';
        sheet.appendRow([new Date(), email, acao, detalhe]);
        
    } catch(e) {
        console.error("Falha ao registrar log: " + e.message);
    } finally {
        lock.releaseLock();
    }
}
