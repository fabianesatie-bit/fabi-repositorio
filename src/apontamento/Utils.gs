/**
 * UTILITÁRIOS E MIDDLEWARE ARQUITETURAL
 */

/**
 * REGRA GLOBAL DE NORMALIZAÇÃO DE FILIAIS (MIDDLEWARE > 3000)
 * Se o número da filial for estritamente superior a 3000, subtrai 3000 em memória.
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Normaliza textos para comparação sem acento e em caixa alta.
 */
function normalizarTextoUpper(texto) {
  if (!texto) return '';
  return texto.toString().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

/**
 * Converte valores variados para número seguro (formato BR ou float).
 */
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
