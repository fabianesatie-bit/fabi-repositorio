/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS - COMPLETO E BLINDADO
 * Atualizado em: 03 de Outubro de 2026
 * Arquivo: Config.gs (Concentra variáveis de ambiente e configurações iniciais)
 */

const CONFIG = {
  nomeAbaBase: 'Extração 1', 
  nomeAbaFiliais: 'Filiais',
  nomeAbaFolgas: 'Folgas Agendadas',
  nomeAbaDadosAtivos: 'DADOS_ATIVOS',
  urlPortal: 'https://script.google.com/a/macros/magazineluiza.com.br/s/AKfycbzWpN0-LZa_Lvd6VugBL5k3tWbHa5jdwKhJyD4UCH4Ml00BeNzved493zspKCgZAmz5Vg/exec',
  emailRemetente: 'gplojas@magazineluiza.com.br'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Banco de Horas')
    .addItem('Configurar Gatilhos (Automático)', 'configurarGatilhos')
    .addToUi();
}

/**
 * ROTEADOR DO WEB APP (PORTAL)
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal - Banco de Horas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
