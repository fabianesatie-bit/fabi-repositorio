// =====================================================================
// CONFIGURAÇÕES GLOBAIS E ROTEADOR - ECOSSISTEMA GP360
// MÓDULO: FORMULÁRIO DE INSCRIÇÃO (PROJETO: Base Posicionamento estrategico)
// =====================================================================

// Planilha Base (Consulta de Ativos e Lojas - DADOS_ATIVOS / DADOS_LOJAS / DADOS_ESCRITORIO)
const SPREADSHEET_BASE_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; 

// Planilha de Respostas (Detalhamento Regional)
const SPREADSHEET_DESTINO_ID = '1K9qtR28oeacW9U1prx24NhqOz_p89dabY0rn3mBHIV8'; 

// E-mails Oficiais de Suporte e Gestão
const EMAIL_ROSA = 'rosa.santos@magazineluiza.com.br';
const EMAIL_AGENCIA = 'posicionamentomagalu@ltnbrasil.com.br';

// Super Admins: Usuários com privilégios de gestão
const SUPER_ADMINS = [
  'fabiane.satie@magazineluiza.com.br', 
  'rosa.santos@magazineluiza.com.br',
  'wylmer.garcia@magazineluiza.com.br',
  'maria.antunes@magazineluiza.com.br'
];

/**
 * Função Roteadora Principal (doGet)
 * Serve a interface do formulário (Index.html) de forma limpa e ininterrupta.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inscrição - Posicionamento Estratégico 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
