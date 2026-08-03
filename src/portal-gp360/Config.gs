// =============================================================================
// CONFIGURAÇÕES GLOBAIS DO BANCO DE DADOS GP - VERSÃO OTIMIZADA
// =============================================================================

const SPREADSHEET_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; // ID Planilha GP 360 (DB_MASTER)
const EVIDENCIAS_FOLDER_ID = '1v28G-ZDd6yQpjTUvBNcODlpxkM5AeaQZ'; // ID Pasta do Drive para Evidências
const SPREADSHEET_DASH = '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc'; // Base de Indicadores Geral (DB_DASH)
const SPREADSHEET_LOG_ID = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY'; // ID Planilha Auditoria Exclusiva

// Lista de administradores com acesso irrestrito
const SUPER_ADMINS_EMAILS = [
  "fabiane.satie@magazineluiza.com.br",
  "gplojas@magazineluiza.com.br",
  "tarcisio.maniglia@magazineluiza.com.br"
];

/**
 * Renderiza a página principal do Portal GP 360 injetando os componentes HTML.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal GP 360')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}
