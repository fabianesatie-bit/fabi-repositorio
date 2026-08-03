// =============================================================================
// MOTOR DE SEGURANÇA: CONTROLE DE ACESSO E CONTEXTO ATIVO (SERVER-SIDE)
// =============================================================================

/**
 * Retorna o perfil de permissões do usuário logado (RBAC).
 */
function obterControleAcesso(email) {
  if (!email) return { autorizado: false, erro: "Sessão não identificada." };

  const emailNorm = email.toLowerCase().trim();
  const isSuper = SUPER_ADMINS_EMAILS.includes(emailNorm);

  if (isSuper) {
    return {
      autorizado: true,
      email: emailNorm,
      nome: emailNorm === "fabiane.satie@magazineluiza.com.br" ? "FABIANE SATIE" : "SUPER ADMIN GP",
      cargo: "Administrador",
      isSuperAdmin: true,
      isAdmin: true,
      isConfigAdmin: true,
      isGerenteGP: true,
      regionais: ["TODAS"],
      diretoriasAtendidas: ["TODAS"]
    };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetUsuarios = ss.getSheetByName('DADOS_USUARIOS');
    if (!sheetUsuarios) return { autorizado: false, erro: "Base de usuários inacessível." };
    const dadosUsuarios = sheetUsuarios.getDataRange().getValues();
    for (let i = 1; i < dadosUsuarios.length; i++) {
      const emailLinha = dadosUsuarios[i][0] ? String(dadosUsuarios[i][0]).toLowerCase().trim() : "";
      if (emailLinha === emailNorm) {
        const status = dadosUsuarios[i][8] ? String(dadosUsuarios[i][8]).toLowerCase().trim() : "ativo";
        if (status === "inativo" || status === "desligado") {
          return { autorizado: false, erro: "Sessão inativa ou revogada." };
        }
        const cargo = dadosUsuarios[i][2] || "Cargo não definido";
        const nivelAcesso = dadosUsuarios[i][5] ? String(dadosUsuarios[i][5]).trim() : "Sem acesso";
        if (nivelAcesso.toLowerCase() === "sem acesso") {
          return { autorizado: false, erro: "Usuário sem permissão de acesso ao Portal." };
        }
        const adminRoles = ["GerenteGP", "Administrador", "DiretorRH"];
        const isAdmin = adminRoles.includes(cargo) || adminRoles.includes(nivelAcesso);
        const isGerenteGP = (cargo === "GerenteGP" || nivelAcesso === "GerenteGP" || isAdmin);

        return {
          autorizado: true,
          email: emailNorm,
          nome: dadosUsuarios[i][1] || "Nome não cadastrado",
          cargo: cargo,
          isSuperAdmin: (cargo === "Administrador" || nivelAcesso === "Administrador"),
          isAdmin: isAdmin,
          isConfigAdmin: isAdmin,
          isGerenteGP: isGerenteGP,
          regionais: dadosUsuarios[i][4] ? String(dadosUsuarios[i][4]).split(',').map(r => r.trim()) : [],
          diretoriasAtendidas: dadosUsuarios[i][3] ? String(dadosUsuarios[i][3]).split(',').map(d => d.trim()) : []
        };
      }
    }
  } catch (e) {
    return { autorizado: false, erro: "Falha de comunicação com o servidor de acessos." };
  }
  return { autorizado: false, erro: "Usuário não localizado na base corporativa." };
}

/**
 * Valida se o usuário tem permissão para visualizar dados de uma filial específica (RLS).
 */
function validarAcessoFilial(email, filialId) {
  const controle = obterControleAcesso(email);
  if (!controle.autorizado) return { autorizado: false };
  if (controle.isSuperAdmin) return { autorizado: true, controle: controle };
  try {
    let targetFilial = parseInt(normalizarFilialId(filialId), 10);
    if (isNaN(targetFilial)) return { autorizado: false };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLojas = ss.getSheetByName('DADOS_LOJAS');
    if (!sheetLojas) return { autorizado: false };
    const dadosLojas = sheetLojas.getDataRange().getValues();
    for (let i = 1; i < dadosLojas.length; i++) {
      let idLoja = parseInt(normalizarFilialId(dadosLojas[i][0]), 10);

      if (idLoja === targetFilial) {
        const regionalLoja = dadosLojas[i][2] ? String(dadosLojas[i][2]).trim() : "";
        const autorizado = controle.regionais.includes(regionalLoja);
        return { autorizado: autorizado, controle: controle };
      }
    }
  } catch (e) {}
  return { autorizado: false };
}
