/**
 * PONTO DE ENTRADA E ROTEAMENTO HTTP - ECOSSISTEMA GP360
 */

function doGet(e) {
  // Verifica se a requisição é do Portal do Gerente via parâmetro da URL
  if (e && e.parameter && (e.parameter.mode === 'certificado' || e.parameter.filialId || e.parameter.filial)) {
    var templatePortal = HtmlService.createTemplateFromFile('FormCertificado');
    templatePortal.filialId = e.parameter.filialId || e.parameter.filial || '';
    templatePortal.chapa = e.parameter.chapa || '';
    
    return templatePortal.evaluate()
      .setTitle('Portal do Gerente - Inclusão de Certificados | GP360')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Acesso Padrão ao Painel Executivo de Jornada
  var templateMain = HtmlService.createTemplateFromFile('Index');
  return templateMain.evaluate()
    .setTitle('Apontamento - Operações de Loja | GP360')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * RETORNA A URL OFICIAL PÚBLICA DO WEB APP (/exec)
 */
function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();
  // Se a URL do serviço estiver disponível, retorna a URL pública oficial
  if (url && url.indexOf('script.google.com') !== -1) {
    return url;
  }
  return 'https://script.google.com/a/macros/magazineluiza.com.br/s/AKfycbwtc018whIZD34s5Tlx9s8I6Pg3HFJUT068XyUqVynM/exec';
}

/**
 * EXPOSIÇÃO GLOBAL DE FUNÇÕES DO SERVIDOR PARA O CLIENTE
 */
function obterDadosAuditoria(forceRefresh) {
  return Service_Data.obterDadosAuditoria(forceRefresh);
}

function salvarTelefoneGerente(filialId, novoTelefone) {
  return Service_Data.salvarTelefoneGerente(filialId, novoTelefone);
}

function obterColaboradoresPendentesFilial(filialId) {
  return Service_Data.obterColaboradoresPendentesFilial(filialId);
}

function uploadCertificadoFile(payload) {
  return Service_Data.uploadCertificadoFile(payload);
}

function saveCertificadoTreinamento(payload) {
  return Service_Data.saveCertificadoTreinamento(payload);
}

function sendEmailFilial(filialId, mensagem) {
  return { success: true, message: 'Plano de Ação enviado com sucesso para a filial ' + filialId };
}

function sendEmailRegional(regionalNome) {
  return { success: true, message: 'Notificação em massa enviada para os gerentes da regional ' + regionalNome };
}
