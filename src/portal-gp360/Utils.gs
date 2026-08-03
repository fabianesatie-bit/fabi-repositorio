// =============================================================================
// MIDDLEWARE E FUNÇÕES UTILITÁRIAS DE SUPORTE
// =============================================================================

/**
 * Regra global de higienização de filiais (Séries 3000/4000 para Série Original).
 * @param {string|number} val ID da filial bruto.
 * @return {string} ID da filial higienizado.
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Motor centralizado de auditoria para registros de acesso e operações.
 */
function registrarAuditoria(evento, detalhe) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ssLog = SpreadsheetApp.openById(SPREADSHEET_LOG_ID);
    const sheetLog = ssLog.getSheetByName('AUDITORIA') || ssLog.getSheets()[0];
    sheetLog.appendRow([new Date(), emailLogado, evento, detalhe]);
  } catch(e) { /* Falha silenciosa para estabilidade da UX */ }
}

/**
 * Extrai o ID limpo de um arquivo do Google Drive a partir de link ou string bruta.
 */
function extrairIdDrive(linkOuId) {
  if (!linkOuId) return "";
  const match = linkOuId.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : linkOuId;
}

/**
 * Formata datas com segurança para o padrão dd/MM/yyyy.
 */
function formatarDataSegura(dataValor) {
  if (!dataValor) return "";
  try { 
    if (dataValor instanceof Date) {
        return Utilities.formatDate(dataValor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) return ds;

    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); 
    }
    return ds; 
  } catch(e) { return String(dataValor).split(' ')[0]; }
}

/**
 * Converte valores de data para timestamp numérico com segurança.
 */
function obterDataRawSegura(dataValor) {
  if (!dataValor) return 0;
  try { 
    if (dataValor instanceof Date) return dataValor.getTime();

    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) {
        let p = ds.split('/');
        if(p.length === 3) return new Date(p[2], p[1]-1, p[0]).getTime();
    } else if (ds.includes('-')) {
        let p = ds.split('-');
        if(p.length === 3) return new Date(p[0], p[1]-1, p[2]).getTime();
    }
    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) return dt.getTime(); 
    return 0; 
  } catch(e) { return 0; }
}
