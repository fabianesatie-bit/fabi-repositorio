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
  if (pageReq === 'gerente') return true; 

  if (userObj.role === 'BLOQUEADO' || userObj.role === 'GERENTE_LOJA' || userObj.role === 'DESCONHECIDO') {
    return false;
  }
  return true;
}

function verificarEhAdminMaster(email) {
  if (!email) return false;
  const user = obterUsuarioLogado(email);
  return user.role === 'MASTER';
}

function obterUsuarioLogado(email) {
  if (!email) return { email: '', nome: 'Visitante', role: 'BLOQUEADO' };
  email = email.toLowerCase().trim();
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

  // Regra Chumbada para Compliance
  if (email === 'angelica@magazineluiza.com.br' || email === 'tarsila.mendonca@magazineluiza.com.br') {
    userObj.role = 'COMPLIANCE';
    userObj.cargo = 'Compliance';
    userObj.nivelAcesso = 'ativo';
    return userObj;
  }

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (dataUsuarios && dataUsuarios.length > 1) {
      
      // Utilização de índices hardcoded da imagem DADOS_USUARIOS para evitar falha de leitura
      const COL_EMAIL = 0;
      const COL_NOME = 1;
      const COL_CARGO = 2;
      const COL_DIR = 3;
      const COL_REG = 4;
      const COL_NIVEL = 5;

      for (let i = 1; i < dataUsuarios.length; i++) {
        const row = dataUsuarios[i];
        const emailCadastrado = row[COL_EMAIL] ? row[COL_EMAIL].toString().toLowerCase().trim() : '';
        
        if (emailCadastrado === email) {
          userObj.nome = row[COL_NOME] ? row[COL_NOME].toString().trim() : apelido;
          userObj.cargo = row[COL_CARGO] ? normalizarTexto(row[COL_CARGO]) : '';
          userObj.diretoria = row[COL_DIR] ? normalizarTexto(row[COL_DIR]) : '';
          userObj.regionais = row[COL_REG] ? normalizarTexto(row[COL_REG]) : '';
          userObj.nivelAcesso = row[COL_NIVEL] ? normalizarTexto(row[COL_NIVEL]) : '';

          // 1. AVALIAÇÃO DE BLOQUEIO ABSOLUTO (Suprema)
          if (userObj.nivelAcesso.includes('bloqueado') || userObj.nivelAcesso.includes('sem acesso') || userObj.nivelAcesso.includes('inativo')) {
            userObj.role = 'BLOQUEADO';
          } 
          // 2. AVALIAÇÃO MASTER
          else if (userObj.nivelAcesso.includes('admin') || userObj.nivelAcesso.includes('master') || email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br') {
            userObj.role = 'MASTER';
          } 
          // 3. AVALIAÇÃO LIDERANÇA (Diretores e Regionais OP)
          else if (userObj.cargo.includes('diretor op') || userObj.cargo.includes('diretoria') || userObj.cargo.includes('regional op')) {
            userObj.role = 'LIDERANCA';
          } 
          // 4. AVALIAÇÃO COMPLIANCE GENÉRICO
          else if (userObj.cargo.includes('compliance') || userObj.cargo.includes('etica')) {
            userObj.role = 'COMPLIANCE';
          } 
          // 5. AVALIAÇÃO GERENTE DE LOJA (Não acessa painel principal)
          else if (userObj.cargo.includes('gerente')) {
            userObj.role = 'GERENTE_LOJA';
          } 
          // 6. EQUIPE GP
          else {
            userObj.role = 'GP';
          }
          break; // Usuário encontrado
        }
      }
    }
  } catch (err) {
    Logger.log("Erro na validação do usuário: " + err.message);
  }

  // Fallback para Fabiane se ela for removida da planilha acidentalmente
  if (userObj.role === 'DESCONHECIDO' && (email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br')) {
     userObj.role = 'MASTER';
  }

  return userObj;
}
