/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Config.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

// IDs de Planilhas e BDD Master
var SPREADSHEET_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; // DB_MASTER
var SPREADSHEET_DASH = '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc'; // DB_DASH
var SPREADSHEET_LOG_ID = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY'; // DB_LOG

// IDs de Planilhas Externas Integradas
var SPREADSHEET_SOCIAL_ID = '1InLKT3qmWxAv7N-U1tyoSW0tNI-Qc4LTW0vyza1oSg0'; // Atendimento Social
var SPREADSHEET_APURACOES_ID = '1tn2FiNVWVMFM-3DC14_L0LaHGmd9O-teU2cvsIoajkk'; // Apurações / Feedback / Desligamentos

// Pastas do Google Drive
var EVIDENCIAS_FOLDER_ID = '1v28G-ZDd6yQpjTUvBNcODlpxkM5AeaQZ';

// Whitelist de Cargos Autorizados ao Portal GP360
var PERMITTED_ROLES = ['Administrador', 'GERENTERH', 'Coordenador'];

// Meta de Gamificação - Trilha da Montanha (Everest Reajustada para 100 Moedas)
var META_EVEREST = 100;

/**
 * Ponto de entrada da aplicação Web App com suporte PWA
 */
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Portal GP 360°')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função utilitária para inclusão de arquivos HTML no template
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
