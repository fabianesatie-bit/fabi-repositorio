/**
 * UTILITÁRIOS E MIDDLEWARE ARQUITETURAL
 */

function normalizarFilialId(val) {
  if (val === null || val === undefined || val === '') return "";
  var str = String(val).trim();
  var digits = str.replace(/\D/g, '');
  if (digits.length > 0) {
    var num = parseInt(digits, 10);
    if (!isNaN(num)) return String(num);
  }
  return str;
}

function normalizarTextoUpper(texto) {
  if (!texto) return '';
  return texto.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function parseFloatBR(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = String(val).replace(/[^0-9.,-]/g, '').trim();
  if (str === '') return 0;
  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

var Utils = {
  normalizarFilialId: normalizarFilialId,
  normalizarTextoUpper: normalizarTextoUpper,
  parseFloatBR: parseFloatBR
};
