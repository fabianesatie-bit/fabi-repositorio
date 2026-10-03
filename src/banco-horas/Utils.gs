/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS
 * Arquivo: Utils.gs (Funções de apoio, formatação e Middlewares Globais)
 */

/**
 * Função Padrão OBRIGATÓRIA - Normaliza o ID da filial aplicando a regra > 3000
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

function normalizarTextoSimples(texto) {
  if (!texto) return '';
  return texto.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

/**
 * Encontra o índice da coluna dinamicamente pelo nome do cabeçalho (BLINDADO CONTRA FALSOS POSITIVOS)
 */
function obterIndice(headers, termosBusca, indicePadrao) {
  if (!headers || !headers.length) return indicePadrao;
  for (let i = 0; i < headers.length; i++) {
    const h = normalizarTextoSimples(headers[i]);
    for (let j = 0; j < termosBusca.length; j++) {
      let termo = normalizarTextoSimples(termosBusca[j]);
      
      // Trava de segurança: Se procurar por "id", exige correspondência exata para não casar com "unidade" ou "validade"
      if (termo === "id") {
        if (h === "id" || h === "id_colaborador" || h === "id colaborador") {
          return i;
        }
      } else {
        if (h.includes(termo)) return i;
      }
    }
  }
  return indicePadrao;
}

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
 * Leitura e conversão de Horas (decimais com vírgula ou formato HH:mm)
 */
function parseHoras(valor) {
  if (valor === "" || valor === null || valor === undefined) return 0;
  
  if (valor instanceof Date) {
    return Number((valor.getHours() + (valor.getMinutes() / 60)).toFixed(2));
  }
  
  if (typeof valor === 'number') return Number(valor.toFixed(2));
  
  let strVal = valor.toString().trim();
  
  if (strVal.includes(':')) {
    let partes = strVal.split(':');
    let h = parseFloat(partes[0]) || 0;
    let m = parseFloat(partes[1]) || 0;
    return Number((h + (m / 60)).toFixed(2));
  }
  
  if (strVal.includes(',')) {
    strVal = strVal.replace(/\./g, '').replace(',', '.');
  }
  let num = parseFloat(strVal);
  return isNaN(num) ? 0 : Number(num.toFixed(2));
}

function isEmAtividadeNormal(situacao) {
  if (!situacao) return true; // Se não tem apontamento negativo, considera ativo
  let s = situacao.toString().trim().toUpperCase();
  return s === "" || s === "EM ATIVIDADE NORMAL" || s === "ATIVO" || s.includes("ATIVIDADE NORMAL");
}
