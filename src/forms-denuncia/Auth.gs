// =============================================================================
// AUTENTICAÇÃO, CONTROLE DE ACESSO E AUTORIZAÇÃO
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: Auth.gs
// =============================================================================

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
    autorizado = verificarAutorizacao(email, pageReq);
    usuario = obterUsuarioLogado(email);
  } catch (err) {
    Logger.log("Erro na validação de sessão: " + err.message);
  }

  return { email: email, autorizado: autorizado, usuario: usuario };
}

function verificarAutorizacao(email, pageReq) {
  if (!email) return false;
  email = email.toLowerCase().trim();

  if (pageReq === 'gerente') {
    return true;
  }

  if (email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br') {
    return true;
  }

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (!dataUsuarios) return false;

    const headers = dataUsuarios[0];
    const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
    const colHierarquia = obterIndiceCabecalho(headers, ['cargo'], 2);
    const colNivel = obterIndiceCabecalho(headers, ['nivel_acesso'], 5);

    for (let i = 1; i < dataUsuarios.length; i++) {
      const emailCadastrado = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
      if (emailCadastrado === email) {
        const hierarquia = dataUsuarios[i][colHierarquia] ? normalizarTexto(dataUsuarios[i][colHierarquia]) : '';
        const nivel = dataUsuarios[i][colNivel] ? normalizarTexto(dataUsuarios[i][colNivel]) : '';

        // BLOQUEIO RIGOROSO: Checa níveis restritivos antes de aprovar por cargo
        if (nivel.includes('bloqueado') || nivel.includes('inativo') || nivel.includes('sem acesso')) {
          return false;
        }

        if (hierarquia.includes('diretor op') || hierarquia.includes('diretoria op') || hierarquia.includes('gerente regional op') || hierarquia.includes('regional op')) {
          return true;
        }

        if (hierarquia.includes('gp') || hierarquia.includes('coord') || hierarquia.includes('diretor')) {
          return true;
        }

        if (/gerente.*loja/i.test(hierarquia) || hierarquia === 'gerente') return false;
        
        return true; 
      }
    }
  } catch (err) {}
  return false;
}

function verificarEhAdminMaster(email) {
  if (!email) return false;
  email = email.toLowerCase().trim();

  if (email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br') {
    return true;
  }

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (!dataUsuarios) return false;

    const headers = dataUsuarios[0];
    const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
    const colNivel = obterIndiceCabecalho(headers, ['nivel_acesso'], 5);

    for (let i = 1; i < dataUsuarios.length; i++) {
      const emailCadastrado = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
      if (emailCadastrado === email) {
        const nivel = dataUsuarios[i][colNivel] ? normalizarTexto(dataUsuarios[i][colNivel]) : '';
        if (nivel.includes('admin') || nivel.includes('master')) {
          return true;
        }
      }
    }
  } catch (e) {
    Logger.log("Erro na checagem de Admin Master: " + e.message);
  }
  return false;
}

function obterUsuarioLogado(email) {
  if (!email) return { email: '', nome: 'COORDENADOR GP', apelido: 'GP', cargo: 'GP', ehRegionalOuDiretor: false };
  let apelido = email.split('@')[0];
  if (apelido.includes('.')) {
    const parts = apelido.split('.');
    apelido = parts[0].charAt(0).toUpperCase() + parts[0].slice(1);
  } else {
    apelido = apelido.charAt(0).toUpperCase() + apelido.slice(1);
  }

  let cargoCadastrado = 'GP';
  let ehRegionalOuDiretor = false;

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (dataUsuarios && dataUsuarios.length > 1) {
      const headers = dataUsuarios[0];
      const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
      const colCargo = obterIndiceCabecalho(headers, ['cargo'], 2);

      for (let i = 1; i < dataUsuarios.length; i++) {
        const emailCad = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
        if (emailCad === email.toLowerCase().trim()) {
          cargoCadastrado = dataUsuarios[i][colCargo] ? dataUsuarios[i][colCargo].toString().trim() : 'GP';
          const cargoNorm = normalizarTexto(cargoCadastrado);
          if (cargoNorm.includes('gerente regional op') || cargoNorm.includes('regional op') || cargoNorm.includes('diretor op') || cargoNorm.includes('diretoria op')) {
            ehRegionalOuDiretor = true;
          }
          break;
        }
      }
    }
  } catch (e) {}

  return {
    email: email,
    nome: email.split('@')[0].replace('.', ' ').toUpperCase(),
    apelido: apelido,
    cargo: cargoCadastrado,
    ehRegionalOuDiretor: ehRegionalOuDiretor
  };
}
