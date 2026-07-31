// =============================================================================
// GESTÃO DO DICIONÁRIO DE PRÊMIOS, NATUREZAS E AVISOS (GERENTEGP CRUD)
// =============================================================================

function getDicionarioPremios() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');

    if (!abaPremios) {
      abaPremios = ss.insertSheet('DICIONARIO_PREMIOS');
      abaPremios.appendRow(['ID_Premio', 'Nome_Premio', 'Custo_ou_Meta', 'Icone']);
      abaPremios.getRange(1, 1, 1, 4).setFontWeight('bold');
    }

    const dados = abaPremios.getDataRange().getValues();
    let lista = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) {
        lista.push({
          id: String(dados[i][0]),
          nome: String(dados[i][1] || ""),
          meta: Number(dados[i][2]) || 0,
          icone: String(dados[i][3] || "🏆")
        });
      }
    }
    return lista;
  } catch (e) {
    return [];
  }
}

function salvarPremio(premio) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.isGerenteGP && !controle.isSuperAdmin) {
      return { sucesso: false, erro: "Acesso negado: Perfil sem permissão de GerenteGP." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
    if (!abaPremios) {
      abaPremios = ss.insertSheet('DICIONARIO_PREMIOS');
      abaPremios.appendRow(['ID_Premio', 'Nome_Premio', 'Custo_ou_Meta', 'Icone']);
    }

    const novoId = premio.id || 'PRM-' + new Date().getTime();
    const dados = abaPremios.getDataRange().getValues();
    let linhaExistente = -1;

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(novoId)) {
        linhaExistente = i + 1;
        break;
      }
    }

    if (linhaExistente > 0) {
      abaPremios.getRange(linhaExistente, 2).setValue(premio.nome);
      abaPremios.getRange(linhaExistente, 3).setValue(premio.meta);
      abaPremios.getRange(linhaExistente, 4).setValue(premio.icone);
    } else {
      abaPremios.appendRow([novoId, premio.nome, premio.meta, premio.icone]);
    }

    return { sucesso: true, premios: getDicionarioPremios() };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

function deletarPremio(idPremio) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.isGerenteGP && !controle.isSuperAdmin) {
      return { sucesso: false, erro: "Acesso negado." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
    if (!abaPremios) return { sucesso: false, erro: "Aba não encontrada." };

    const dados = abaPremios.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idPremio)) {
        abaPremios.deleteRow(i + 1);
        return { sucesso: true, premios: getDicionarioPremios() };
      }
    }

    return { sucesso: false, erro: "Prêmio não encontrado." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

function adicionarNatureza(novaNatureza) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Apenas administradores do sistema podem registrar naturezas." };
  }
  try {
    if(!novaNatureza) return { erro: "Nome inválido!" };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
    const valores = sheetConfig.getDataRange().getValues().map(r => String(r[0]).toLowerCase().trim());
    if (valores.includes(novaNatureza.toLowerCase().trim())) return { erro: "Esta natureza já está cadastrada!" };

    sheetConfig.appendRow([novaNatureza.trim()]);
    return { sucesso: true, naturezas: carregarNaturezasSeguras(ss) };
  } catch (e) { return { erro: e.message }; }
}

function excluirNatureza(naturezaTexto) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Operação restrita a administradores." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
    const valores = sheetConfig.getDataRange().getValues();
    for (let i = 1; i < valores.length; i++) {
      if (String(valores[i][0]).trim() === naturezaTexto.trim()) {
        sheetConfig.deleteRow(i + 1);
        return { sucesso: true, naturezas: carregarNaturezasSeguras(ss) };
      }
    }
    return { erro: "Natureza não localizada." };
  } catch (e) { return { erro: e.message }; }
}

function publicarAviso(mensagem, diretoriaAlvo) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Publicação restrita a diretores ou administradores." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    let nomeAutor = controle.nome;
    const usu = ss.getSheetByName('DADOS_USUARIOS').getDataRange().getValues();
    let emails = [];

    let dirAlvoNorm = String(diretoriaAlvo).toLowerCase().trim();
    for (let i = 1; i < usu.length; i++) {
       let em = String(usu[i][0]).toLowerCase().trim();
       let dirColNormArray = String(usu[i][3]).toLowerCase().split(',').map(s => s.trim());

       if(em && (dirAlvoNorm === "todos" || dirColNormArray.includes(dirAlvoNorm))) {
           emails.push(em);
       }
    }
    let prefixo = diretoriaAlvo === "Todos" ? "" : `[${diretoriaAlvo}] `;
    sheetAvisos.appendRow([new Date(), nomeAutor, prefixo + mensagem]);
    if (emails.length > 0) {
      MailApp.sendEmail({
        bcc: emails.join(','), subject: "📢 Novo Recado no Portal GP 360", 
        htmlBody: `<div style="font-family:Arial;border:1px solid #ddd;padding:20px;border-radius:8px;"><h2 style="color:#0086ff;">Novo Recado!</h2><p><strong>De:</strong> ${nomeAutor}</p><p><strong>Para:</strong> ${diretoriaAlvo}</p><div style="background:#fff3cd;padding:15px;border-left:5px solid #ffcc00;">"${mensagem}"</div><br><a href="${ScriptApp.getService().getUrl()}" style="background:#0086ff;color:white;padding:10px 20px;text-decoration:none;border-radius:5px;">Acessar Portal</a></div>`
      });
    }
    return { sucesso: true };
  } catch (e) { return { erro: e.message }; }
}

function excluirAvisoPlanilha(dataAviso, mensagemAviso) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Apenas administradores podem excluir recados." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    const dados = sheetAvisos.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
       if (formatarDataSegura(dados[i][0]) === dataAviso && String(dados[i][2]) === mensagemAviso) {
           sheetAvisos.deleteRow(i + 1); return { sucesso: true };
       }
    }
    return { erro: "Recado não localizado." };
  } catch (e) { return { erro: e.message }; }
}
