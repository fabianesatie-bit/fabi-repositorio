// =============================================================================
// GALERIA DE EVIDÊNCIAS E BYPASS CORS BASE64 (HIERARQUIA BLINDADA)
// =============================================================================

function getEvidenciasUsuario() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLancamentos = ss.getSheetByName('DADOS_LANCAMENTOS');
    const sheetUsu = ss.getSheetByName('DADOS_USUARIOS');
    if (!sheetLancamentos || !sheetUsu) return [];

    let mapUsu = {};
    const dUsu = sheetUsu.getDataRange().getValues();
    for(let i=1; i<dUsu.length; i++){
        let em = String(dUsu[i][0]).toLowerCase().trim();
        let dirArr = dUsu[i][3] ? String(dUsu[i][3]).toLowerCase().split(',').map(s=>s.trim()) : [];
        let regArr = dUsu[i][4] ? String(dUsu[i][4]).toLowerCase().split(',').map(s=>s.trim()) : [];
        mapUsu[em] = { dir: dirArr, reg: regArr };
    }

    const dados = sheetLancamentos.getDataRange().getValues();
    let evidencias = [];

    let myDir = controle.diretoriasAtendidas ? controle.diretoriasAtendidas.map(d=>d.toLowerCase().trim()) : [];
    let myReg = controle.regionais ? controle.regionais.map(r=>r.toLowerCase().trim()) : [];

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const emailAutor = String(linha[2] || "").toLowerCase().trim();
      const linkEvidencia = String(linha[11] || "").trim();
      if (!linkEvidencia) continue;

      let hasAccess = false;
      if (controle.isSuperAdmin) {
          hasAccess = true; 
      } else if (controle.isGerenteGP) {
          if (emailAutor === emailLogado) {
              hasAccess = true;
          } else if (mapUsu[emailAutor]) {
              let overlapDir = mapUsu[emailAutor].dir.some(d => myDir.includes(d));
              let overlapReg = mapUsu[emailAutor].reg.some(r => myReg.includes(r));
              if (overlapDir || overlapReg) hasAccess = true;
          }
      } else {
          if (emailAutor === emailLogado) hasAccess = true; 
      }

      if (hasAccess) {
        const match = linkEvidencia.match(/id=([a-zA-Z0-9_-]+)/) || linkEvidencia.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          const fileId = match[1];
          try {
            let thumbUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`;
            evidencias.push({
              id: fileId,
              nome: "Evidência Visual", // Nome genérico injetado para evitar timeout do DriveApp
              motivo: String(linha[4] || "Ação de Governança"),
              destino: String(linha[3] || "Filial"),
              data: formatarDataSegura(linha[12] || linha[1]),
              thumbUrl: thumbUrl,
              downloadUrl: `https://drive.google.com/uc?export=download&id=${fileId}`
            });
          } catch (e) {
            continue;
          }
        }
      }
    }
    return evidencias.reverse().slice(0, 150); // Limite injetado para performance do DOM
  } catch (e) {
    return [];
  }
}

function getImagemBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);

    return {
      sucesso: true,
      mimeType: blob.getContentType(),
      base64: base64
    };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}
