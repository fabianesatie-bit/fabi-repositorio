/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Drive.gs
 * Subpasta Monorepo: src/portal-gp360/
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
    var dtObj = dados[i][1] || dados[i][12]; // Col B (Data_Hora) ou Col M (Data_Ini)
    var dtParsed = dtObj instanceof Date ? dtObj : new Date(dtObj);
    var mReg = dtParsed.getMonth() + 1;
    var aReg = dtParsed.getFullYear();

    if (mReg === mesAlvo && aReg === anoAlvo) {
      var urlEvidencia = String(dados[i][11] || '').trim(); // Col L: Link_Evidencia (index 11)
      if (urlEvidencia) {
        var fileId = extrairIdDrive(urlEvidencia);
        galeria.push({
          id: dados[i][0], // Col A: ID_Lancamento
          data: formatarDataSegura(dtObj),
          loja: String(dados[i][3] || 'Filial/Regional'), // Col D: Destino_Filial_Regional
          motivo: String(dados[i][4] || 'Atividade'),     // Col E: Motivo_Meta
          url: urlEvidencia,
          fileId: fileId,
          thumbUrl: fileId ? 'https://lh3.googleusercontent.com/d/' + fileId + '=w400' : urlEvidencia
        });
      }
    }
    if (galeria.length >= 150) break;
  }
  return galeria;
}
