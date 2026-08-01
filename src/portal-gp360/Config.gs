/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Config.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

var SPREADSHEET_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; // DB_MASTER
var SPREADSHEET_DASH = '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc'; // DB_DASH
var SPREADSHEET_LOG_ID = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY'; // DB_LOG

// Integracoes Externas de Moedas
var SPREADSHEET_SOCIAL_ID = '1InLKT3qmWxAv7N-U1tyoSW0tNI-Qc4LTW0vyza1oSg0'; // Atendimento Social
var SPREADSHEET_APURACOES_ID = '1tn2FiNVWVMFM-3DC14_L0LaHGmd9O-teU2cvsIoajkk'; // Apurações / Desligamentos

// Pastas
var EVIDENCIAS_FOLDER_ID = '1v28G-ZDd6yQpjTUvBNcODlpxkM5AeaQZ';

// Trava Rigida de Permissão
var PERMITTED_ROLES = ['administrador', 'gerenterh', 'coordenador'];

// Meta de Gamificação
var META_EVEREST = 100;

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Portal GP 360°')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
