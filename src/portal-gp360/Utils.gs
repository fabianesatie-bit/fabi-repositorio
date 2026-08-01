/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Utils.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Extrai o ID alfanumérico do Google Drive de uma URL ou ID direto
 */
function extrairIdDrive(linkOuId) {
  if (!linkOuId) return '';
  var str = String(linkOuId).trim();
  if (str.length === 33 || (str.length >= 25 && !str.includes('/'))) {
    return str;
  }
  var match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

/**
 * Padroniza saída de datas para string dd/MM/yyyy
 */
function formatarDataSegura(dataValor) {
  if (!dataValor) return '';
  try {
    if (dataValor instanceof Date) {
      return Utilities.formatDate(dataValor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    var d = new Date(dataValor);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return String(dataValor);
  } catch (e) {
    return String(dataValor);
  }
}

/**
 * Converte strings ou objetos de data em Timestamp para ordenação
 */
function obterDataRawSegura(dataValor) {
  if (!dataValor) return 0;
  try {
    if (dataValor instanceof Date) return dataValor.getTime();
    var parts = String(dataValor).split('/');
    if (parts.length === 3) {
      var d = new Date(parts[2], parts[1] - 1, parts[0]);
      return d.getTime();
    }
    var d2 = new Date(dataValor);
    return isNaN(d2.getTime()) ? 0 : d2.getTime();
  } catch (e) {
    return 0;
  }
}
