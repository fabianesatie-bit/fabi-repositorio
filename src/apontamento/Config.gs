/**
 * CONFIGURAÇÕES GLOBAIS - MÓDULO AUDITORIA GP360
 * Subpasta Monorepo: src/auditoria/
 */

const SPREADSHEET_DB_MASTER_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA';
const SPREADSHEET_AUDITORIA_ID = '1RHbEZ67n9ZjieKiMcpm_qqTVmFsmjyv9nIfDqpwm0WM';

const TAB_NAMES = {
  USUARIOS: 'DADOS_USUARIOS',
  LOJAS: 'DADOS_LOJAS',
  FORA_JORNADA_PATTERN: 'Acesso fora da jornada',
  HORAS_EXTRAS_PATTERN: 'Horas extras',
  BRITANICOS_PATTERN: 'Ajuste / Britânicos',
  CERTIFICADOS: 'Comprovantes_Treinamento'
};

const EMAIL_ORGANIZADOR_PADRAO = 'gplojas@magazineluiza.com.br';
const LINK_CURSO_OBRIGATORIO = 'https://universidadeluiza.com.br/app/home/canal/magalu?section=csc-conte-sempre-comigo&trail=ponto-eletronico-o-guia-essencial-do-colaborador';

const SUPER_ADMIN_EMAILS = [
  'fabiane.satie@magazineluiza.com.br',
  'tarcisio.maniglia@magazineluiza.com.br',
  'gplojas@magazineluiza.com.br'
];


Utils.gs

/**
 * UTILITÁRIOS E MIDDLEWARE ARQUITETURAL
 */

function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
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
