/**
 * CONFIGURAÇÕES GLOBAIS - MÓDULO AUDITORIA GP360
 * Subpasta Monorepo: src/auditoria/
 */

var Utils = Utils || {};

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
