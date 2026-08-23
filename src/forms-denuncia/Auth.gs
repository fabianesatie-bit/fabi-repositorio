function doGet(e) {
  let emailLogado = '';
  try {
    emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (err) {}

  let page = e && e.parameter && e.parameter.page ? e.parameter.page : 'inicio';
  let action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  let id = e && e.parameter && e.parameter.id ? e.parameter.id : '';
  let idFeedback = e && e.parameter && e.parameter.idFeedback ? e.parameter.idFeedback : '';
  let idDesligamento = e && e.parameter && e.parameter.idDesligamento ? e.parameter.idDesligamento : '';

  const template = HtmlService.createTemplateFromFile('Index');
  template.emailLogado = emailLogado;
  template.page = page;
  template.action = action;
  template.editId = id;
  template.idFeedback = idFeedback;
  template.idDesligamento = idDesligamento;

  return template.evaluate()
    .setTitle('GP Magalu - Case-Flow & Feedback')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterSessaoInicial(pageReq) {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (e) {}

  let autorizado = false;
  let usuario = null;

  try {
    usuario = obterUsuarioLogado(email);
    autorizado = verificarAutorizacao(usuario, pageReq);
  } catch (err) {
    Logger.log("Erro na validação de sessão: " + err.message);
  }

  return { email: email, autorizado: autorizado, usuario: usuario };
}

function verificarAutorizacao(userObj, pageReq) {
  if (!userObj || !userObj.email) return false;

  // 1. Bloqueio absoluto para inativos ou desconhecidos
  if (userObj.role === 'BLOQUEADO' || userObj.role === 'DESCONHECIDO') {
    return false;
  }

  // 2. Gerente de Loja: SÓ ACESSA se a página solicitada for 'gerente' (Link do Feedback)
  if (userObj.role === 'GERENTE_LOJA') {
    return pageReq === 'gerente';
  }

  // 3. Perfis Autorizados (GP, MASTER, LIDERANCA, COMPLIANCE): Acesso Liberado ao Painel
  return true;
}

function verificarEhAdminMaster(email) {
  if (!email) return false;
  const user = obterUsuarioLogado(email);
  return user.role === 'MASTER';
}

function obterUsuarioLogado(email) {
  if (!email) return { email: '', nome: 'Visitante', role: 'BLOQUEADO' };
  
  const emailBuscado = typeof normalizarTexto === 'function' ? normalizarTexto(email) : email.toLowerCase().trim();
  let apelido = email.split('@')[0].split('.')[0];
  apelido = apelido.charAt(0).toUpperCase() + apelido.slice(1);

  let userObj = {
    email: email,
    nome: apelido,
    apelido: apelido,
    cargo: 'Não Cadastrado',
    role: 'DESCONHECIDO',
    regionais: '',
    diretoria: '',
    nivelAcesso: 'sem acesso'
  };

  // 1. GERENTE DE LOJA: Se o e-mail começa ou contém 'gerente' (ex: gerente001@magazineluiza.com.br)
  if (emailBuscado.startsWith('gerente') || emailBuscado.includes('gerente')) {
    userObj.role = 'GERENTE_LOJA';
    userObj.cargo = 'Gerente de Loja';
    userObj.nome = 'GERENTE DE LOJA';
    return userObj;
  }

  // 2. COMPLIANCE (Exceção chumbada)
  if (emailBuscado.includes('angelica') || emailBuscado.includes('tarsila')) {
    userObj.role = 'COMPLIANCE';
    userObj.cargo = 'Compliance';
    userObj.nivelAcesso = 'ativo';
    return userObj;
  }

  // 3. BUSCA NA PLANILHA DADOS_USUARIOS (Para Marcia Almeida, Fabiane, Lorena, Coordenadores, etc.)
  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (dataUsuarios && dataUsuarios.length > 1) {
      
      const COL_EMAIL = 0;
      const COL_NOME = 1;
      const COL_CARGO = 2;
      const COL_DIR = 3;
      const COL_REG = 4;
      const COL_NIVEL = 5;

      for (let i = 1; i < dataUsuarios.length; i++) {
        const row = dataUsuarios[i];
        const emailCadastrado = row[COL_EMAIL] ? (typeof normalizarTexto === 'function' ? normalizarTexto(row[COL_EMAIL]) : row[COL_EMAIL].toString().toLowerCase().trim()) : '';
        
        if (emailCadastrado === emailBuscado) {
          userObj.nome = row[COL_NOME] ? row[COL_NOME].toString().trim() : apelido;
          userObj.cargo = row[COL_CARGO] ? normalizarTexto(row[COL_CARGO]) : '';
          userObj.diretoria = row[COL_DIR] ? normalizarTexto(row[COL_DIR]) : '';
          userObj.regionais = row[COL_REG] ? normalizarTexto(row[COL_REG]) : '';
          userObj.nivelAcesso = row[COL_NIVEL] ? normalizarTexto(row[COL_NIVEL]) : '';

          if (userObj.nivelAcesso.includes('bloqueado') || userObj.nivelAcesso.includes('sem acesso') || userObj.nivelAcesso.includes('inativo')) {
            userObj.role = 'BLOQUEADO';
          } 
          else if (userObj.nivelAcesso.includes('admin') || userObj.nivelAcesso.includes('master') || emailBuscado.includes('fabiane.satie') || emailBuscado.includes('gplojas')) {
            userObj.role = 'MASTER';
          } 
          else if (userObj.cargo.includes('diretor op') || userObj.cargo.includes('diretoria') || userObj.cargo.includes('regional op')) {
            userObj.role = 'LIDERANCA';
          } 
          else if (userObj.cargo.includes('compliance') || userObj.cargo.includes('etica')) {
            userObj.role = 'COMPLIANCE';
          } 
          else {
            userObj.role = 'GP'; // Libera Marcia Almeida e equipe de GP/RH
          }
          break;
        }
      }
    }
  } catch (err) {
    Logger.log("Erro na validação do usuário: " + err.message);
  }

  if (userObj.role === 'DESCONHECIDO' && (emailBuscado.includes('fabiane.satie') || emailBuscado.includes('gplojas'))) {
     userObj.role = 'MASTER';
  }

  return userObj;
}
