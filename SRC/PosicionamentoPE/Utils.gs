/**
 * ARQUIVO: Utils.gs
 * OBJETIVO: Funções de apoio, formatações e middlewares essenciais.
 */

/**
 * REGRA GLOBAL (OBRIGATÓRIA): Normalização de ID de Filiais.
 * Garante que lojas com numeração superior a 3000 retornem à base original (subtrai 3000).
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Converte data do formato do Google Sheets para string DD/MM/AAAA.
 */
function formatarDataParaString(dataObj) {
  if (!dataObj || !(dataObj instanceof Date)) return dataObj;
  var dia = String(dataObj.getDate()).padStart(2, '0');
  var mes = String(dataObj.getMonth() + 1).padStart(2, '0');
  var ano = dataObj.getFullYear();
  return dia + '/' + mes + '/' + ano;
}
