/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Drive.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Recupera os últimos arquivos de evidência salvos para a Galeria com filtro de Mês e Ano
 */
function getEvidenciasUsuario(filtroMes, filtroAno) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) return [];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var aba = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (!aba) return [];

  var agora = new Date();
  var mesAlvo = (filtroMes !== undefined && filtroMes !== null && filtroMes !== '') ? parseInt(filtroMes) : agora.getMonth() + 1;
  var anoAlvo = (filtroAno !== undefined && filtroAno !== null && filtroAno !== '') ? parseInt(filtroAno) : agora.getFullYear();

  var dados = aba.getDataRange().getValues();
  var galeria = [];

  for (var i = dados.length - 1; i >= 1; i--) {
    var dtObj = dados[i][1];
    var dtParsed = dtObj instanceof Date ? dtObj : new Date(dtObj);
    var mReg = dtParsed.getMonth() + 1;
    var aReg = dtParsed.getFullYear();

    if (mReg === mesAlvo && aReg === anoAlvo) {
      var urlEvidencia = String(dados[i][8] || '').trim();
      if (urlEvidencia) {
        var fileId = extrairIdDrive(urlEvidencia);
        galeria.push({
          id: dados[i][0],
          data: formatarDataSegura(dados[i][1]),
          loja: (dados[i][4] || 'Filial') + ' (' + (dados[i][3] || '') + ')',
          motivo: dados[i][5] || 'Atividade',
          url: urlEvidencia,
          fileId: fileId,
          thumbUrl: 'https://lh3.googleusercontent.com/d/' + fileId + '=w400'
        });
      }
    }
    if (galeria.length >= 150) break;
  }
  return galeria;
}
