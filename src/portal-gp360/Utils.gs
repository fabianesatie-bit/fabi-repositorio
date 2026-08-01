/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Utils.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

function normalizarFilialId(id) {
  if (!id && id !== 0) return null;
  var num = parseInt(String(id).replace(/\D/g, ''), 10);
  if (isNaN(num)) return null;
  if (num > 3000) num -= 3000;
  return num;
}

function formatarFilialQuatroDigitos(id) {
  var norm = normalizarFilialId(id);
  if (norm === null) return '0000';
  return ("0000" + norm).slice(-4);
}

function extrairIdDrive(linkOuId) {
  if (!linkOuId) return '';
  var str = String(linkOuId).trim();
  if (str.length === 33 || (str.length >= 25 && !str.includes('/'))) {
    return str;
  }
  var match = str.match(/[-\w]{25,}/);
  return match ? match[0] : str;
}

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
