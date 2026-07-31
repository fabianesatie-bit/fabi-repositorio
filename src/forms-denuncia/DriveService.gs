// =============================================================================
// GERENCIADOR DE ARQUIVOS E BLINDAGEM DE PERMISSÕES NO GOOGLE DRIVE
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: DriveService.gs
// =============================================================================

/**
 * CORREÇÃO DE SEGURANÇA CRÍTICA:
 * Força o arquivo a ser estritamente PRIVADO e remove qualquer herança
 * de compartilhamento público ou corporativo antes de atribuir os acessos específicos.
 */
function aplicarPermissoesArquivo(fileId, emailsString, nivelAcesso) {
  if (!fileId) return;
  try {
    const file = DriveApp.getFileById(fileId);

    // Revoga acesso público por link ou domínio corporativo herdado da pasta
    try {
      file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    } catch (errPriv) {
      Logger.log("Aviso de ajuste de privacidade privada em " + fileId + ": " + errPriv.message);
    }

    if (!emailsString) return;

    const emails = emailsString
      .replace(/;/g, ',')
      .split(',')
      .map(function(e) { return e.trim().toLowerCase(); })
      .filter(function(e) { return e && e.indexOf('@') !== -1; });

    emails.forEach(function(email) {
      try {
        if (nivelAcesso === 'EDITOR') {
          file.addEditor(email);
        } else if (nivelAcesso === 'COMENTADOR') {
          file.addCommenter(email);
        } else {
          file.addViewer(email);
        }
      } catch (errUser) {
        Logger.log("Falha ao conceder acesso ao usuário: " + email + ". Detalhe: " + errUser.message);
      }
    });
  } catch (e) {
    Logger.log("Erro crítico no motor de permissões para o arquivo " + fileId + ": " + e.message);
  }
}

/**
 * Converte e salva evidências em Base64 na pasta corporativa do Drive com privacidade restrita.
 */
function salvarEvidenciasDrive(arquivosBase64, filial) {
  if (!arquivosBase64 || arquivosBase64.length === 0) return '';
  const links = [];
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    arquivosBase64.forEach(function(arq, index) {
      if (arq.dados && arq.dados.indexOf(',') !== -1) {
        const split = arq.dados.split(',');
        const match = split[0].match(/:(.*?);/);
        const contentType = match ? match[1] : 'image/png';
        const rawData = Utilities.base64Decode(split[1]);
        const fileName = "Evidencia_F" + filial + "_" + new Date().getTime() + "_" + (arq.nome || ("Print_" + index + ".png"));
        const file = folder.createFile(Utilities.newBlob(rawData, contentType, fileName));

        // Aplica blindagem de privacidade no arquivo de evidência
        file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        links.push(file.getUrl());
      } else if (arq.dados && arq.dados.indexOf('http') === 0) {
        links.push(arq.dados);
      }
    });
  } catch (e) {
    Logger.log("Erro ao salvar evidências: " + e.message);
  }
  return links.join('\n');
}
