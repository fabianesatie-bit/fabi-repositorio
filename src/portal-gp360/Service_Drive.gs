// =============================================================================
// SERVIÇOS E UTILITÁRIOS DO GOOGLE DRIVE & EVIDÊNCIAS
// =============================================================================

/**
 * Recupera as evidências visuais registradas na base DADOS_LANCAMENTOS.
 * Respeita a jurisdição do usuário (super admin visualiza todas; consultor visualiza suas próprias evidências).
 * 
 * @return {Object} Objeto contendo lista de evidências ou mensagem de erro.
 */
function getEvidenciasUsuario() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      return { erro: "Acesso não autorizado." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!sheet) return { sucesso: true, evidencias: [] };

    const data = sheet.getDataRange().getValues();
    const evidencias = [];

    // Linha 0 é o cabeçalho. Itera do lançamento mais recente ao mais antigo.
    for (let i = data.length - 1; i >= 1; i--) {
      const autorEmail = String(data[i][2]).toLowerCase().trim();
      const linkEvidencia = String(data[i][11] || "").trim();

      if (!linkEvidencia) continue;

      if (controle.isSuperAdmin || autorEmail === emailLogado) {
        const fileId = extrairIdDrive(linkEvidencia);
        evidencias.push({
          id: data[i][0],
          data: formatarDataSegura(data[i][12] || data[i][1]),
          autor: autorEmail,
          destino: data[i][3],
          motivo: data[i][4],
          urlOriginal: linkEvidencia,
          fileId: fileId,
          thumbnailUrl: fileId ? "https://drive.google.com/thumbnail?id=" + fileId : ""
        });
      }

      // Limita a 150 evidências recentes para manter boa performance
      if (evidencias.length >= 150) break;
    }

    return { sucesso: true, evidencias: evidencias };
  } catch (e) {
    return { erro: e.message };
  }
}

/**
 * Converte um arquivo de imagem do Drive em Base64 para exibição segura sem bloqueios de CORS no navegador.
 * 
 * @param {string} fileId ID do arquivo no Google Drive.
 * @return {Object} Objeto contendo o payload Base64 com MIME type ou erro.
 */
function getImagemBase64(fileId) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      return { erro: "Acesso não autorizado." };
    }

    if (!fileId) {
      return { erro: "ID de arquivo inválido ou não informado." };
    }

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const contentType = blob.getContentType();

    return {
      sucesso: true,
      mimeType: contentType,
      base64: "data:" + contentType + ";base64," + base64Data
    };
  } catch (e) {
    return { erro: e.message };
  }
}
