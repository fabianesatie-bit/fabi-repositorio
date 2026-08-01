/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Utils.gs
 * Subpasta Monorepo: src/portal-gp360/
 */


/**
 * Normaliza o ID da filial eliminando duplicações para lojas com número acima de 3000
 * @param {string|number} id - ID bruto da filial
 * @return {string} ID normalizado
 */
function normalizarFilialId(id) {
  if (id === undefined || id === null || id === '') return '';
  var num = parseInt(String(id).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(id).trim();
  if (num > 3000) num -= 3000;
  return String(num);
}

/**
 * Extrai o ID alfanumérico do Google Drive de uma URL ou ID direto
 * @param {string} linkOuId - Link completo do Google Drive ou o ID isolado
 * @return {string} ID extraído
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
 * Padroniza saída de datas para string dd/MM/yyyy tratando problemas de fuso horário (GMT-3)
 * @param {Date|string} dataValor - Data a ser formatada
 * @return {string} Data formatada no padrão dd/MM/yyyy
 */
function formatarDataSegura(dataValor) {
  if (!dataValor) return '';
  try {
    if (dataValor instanceof Date) {
      return Utilities.formatDate(dataValor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }

    var str = String(dataValor).trim();

    // Tratamento direto para formato ISO YYYY-MM-DD sem recuo de fuso horário
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      var partesIso = str.split('T')[0].split('-');
      if (partesIso.length === 3) {
        return partesIso[2] + '/' + partesIso[1] + '/' + partesIso[0];
      }
    }

    // Tratamento direto para formato dd/MM/yyyy
    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(str)) {
      var pBarra = str.split('/');
      var d = ("0" + pBarra[0]).slice(-2);
      var m = ("0" + pBarra[1]).slice(-2);
      var a = pBarra[2].substring(0, 4);
      return d + '/' + m + '/' + a;
    }

    var dObj = new Date(dataValor);
    if (!isNaN(dObj.getTime())) {
      return Utilities.formatDate(dObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return str;
  } catch (e) {
    return String(dataValor);
  }
}


/**
 * Converte strings ou objetos de data em Timestamp para ordenação em arrays
 * @param {Date|string} dataValor - Data de entrada
 * @return {number} Timestamp numérico em milissegundos
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
