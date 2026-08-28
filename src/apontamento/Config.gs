/**
 * CONFIGURAÇÕES GLOBAIS - MÓDULO AUDITORIA GP360
 * Subpasta Monorepo: src/auditoria/
 */

// Planilha Master (Perfis e Usuários)
const SPREADSHEET_DB_MASTER_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA';

// Planilha de Apontamentos de Auditoria
const SPREADSHEET_AUDITORIA_ID = '1RHbEZ67n9ZjieKiMcpm_qqTVmFsmjyv9nIfDqpwm0WM';

// Nomes de referência ou padrões de busca para as abas
const TAB_NAMES = {
  USUARIOS: 'DADOS_USUARIOS',
  FORA_JORNADA_PATTERN: 'Acesso fora da jornada',
  HORAS_EXTRAS_PATTERN: 'Horas extras',
  BRITANICOS_PATTERN: 'Ajuste / Britânicos'
};

// Admins Globais com Acesso Irrestrito
const SUPER_ADMIN_EMAILS = [
  'fabiane.satie@magazineluiza.com.br',
  'tarcisio.maniglia@magazineluiza.com.br',
  'gplojas@magazineluiza.com.br'
];
