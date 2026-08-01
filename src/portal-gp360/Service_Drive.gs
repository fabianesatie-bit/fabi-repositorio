/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Drive.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Recupera os últimos 150 arquivos de evidência salvos para a Galeria
 */
function getEvidenciasUsuario() {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) return [];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (!aba) return [];

  var dados = aba.getDataRange().getValues();
  var galeria = [];

  for (var i = dados.length - 1; i >= 1; i--) {
    var urlEvidencia = String(dados[i][8] || '').trim();
    if (urlEvidencia) {
      var fileId = extrairIdDrive(urlEvidencia);
      galeria.push({
        id: dados[i][0],
        data: formatarDataSegura(dados[i][1]),
        loja: dados[i][4] + ' (' + dados[i][3] + ')',
        motivo: dados[i][5],
        url: urlEvidencia,
        fileId: fileId,
        thumbUrl: 'https://lh3.googleusercontent.com/d/' + fileId + '=w400'
      });
    }
    if (galeria.length >= 150) break;
  }
  return galeria;
}
