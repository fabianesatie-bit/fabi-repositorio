/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Security.gs
 * Subpasta Monorepo: src/portal-gp360/
 */

/**
 * Registra log de auditoria silencioso em planilha exclusiva
 */
function registrarAuditoria(evento, detalhe) {
  try {
    var ssLog = SpreadsheetApp.openById(SPREADSHEET_LOG_ID);
    var abaLog = ssLog.getSheetByName('AUDITORIA') || ssLog.getSheets()[0];
    var email = Session.getActiveUser().getEmail() || 'usuario.desconhecido';
    abaLog.appendRow([new Date(), email, evento, detalhe]);
  } catch (e) {
    Logger.log('Erro ao registrar auditoria: ' + e.toString());
  }
}

/**
 * Obtém perfil e valida nível de acesso do usuário logado contra a aba DADOS_USUARIOS
 */
function obterControleAcesso(email) {
  var emailAtivo = email || Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var abaUsuarios = ss.getSheetByName('DADOS_USUARIOS');
  
  if (!abaUsuarios) {
    return { temAcesso: false, motivo: 'Aba DADOS_USUARIOS não encontrada.' };
  }
  
  var dados = abaUsuarios.getDataRange().getValues();
  var usuarioEncontrado = null;
  
  // Coluna A (0): Email | B (1): Nome | C (2): Cargo | D (3): Diretoria | E (4): Regionais | F (5): Nivel_Acesso | G (6): Foto | H (7): Observações do acesso
  for (var i = 1; i < dados.length; i++) {
    var emailRow = String(dados[i][0]).toLowerCase().trim();
    if (emailRow === String(emailAtivo).toLowerCase().trim()) {
      usuarioEncontrado = {
        email: emailAtivo,
        nome: dados[i][1] || 'Usuário GP',
        cargo: dados[i][2] || '',
        diretoria: dados[i][3] || '',
        regionaisStr: String(dados[i][4] || ''),
        nivelAcesso: String(dados[i][5] || '').trim(),
        fotoUrl: dados[i][6] || '',
        observacoesAcesso: dados[i][7] || ''
      };
      break;
    }
  }

  if (!usuarioEncontrado) {
    registrarAuditoria('BLOQUEIO_ACESSO', 'E-mail ' + emailAtivo + ' não cadastrado em DADOS_USUARIOS.');
    return { temAcesso: false, email: emailAtivo, motivo: 'Usuário não cadastrado na base oficial de acessos.' };
  }

  // Trava de Whitelist Rígida: Apenas Administrador, GERENTERH, Coordenador
  var ehPermitido = PERMITTED_ROLES.some(function(role) {
    return role.toLowerCase() === usuarioEncontrado.nivelAcesso.toLowerCase();
  });

  if (!ehPermitido) {
    registrarAuditoria('BLOQUEIO_PERFIL', 'Acesso negado para ' + emailAtivo + ' com nível: ' + usuarioEncontrado.nivelAcesso);
    return { 
      temAcesso: false, 
      email: emailAtivo, 
      motivo: 'Acesso restrito ao Portal GP 360. Seu nível atual (' + usuarioEncontrado.nivelAcesso + ') não possui permissão neste ambiente.' 
    };
  }

  var isSuperAdmin = (usuarioEncontrado.nivelAcesso.toLowerCase() === 'administrador');
  var isGerenteGP = (usuarioEncontrado.nivelAcesso.toLowerCase() === 'gerenterh');
  var isCoordenador = (usuarioEncontrado.nivelAcesso.toLowerCase() === 'coordenador');

  var regionaisArray = usuarioEncontrado.regionaisStr.split(',')
    .map(function(r) { return r.trim().toUpperCase(); })
    .filter(function(r) { return r.length > 0; });

  return {
    temAcesso: true,
    email: emailAtivo,
    nome: usuarioEncontrado.nome,
    cargo: usuarioEncontrado.cargo,
    diretoria: usuarioEncontrado.diretoria,
    regionais: regionaisArray,
    nivelAcesso: usuarioEncontrado.nivelAcesso,
    fotoUrl: usuarioEncontrado.fotoUrl,
    observacoesAcesso: usuarioEncontrado.observacoesAcesso,
    isSuperAdmin: isSuperAdmin,
    isAdmin: isSuperAdmin || isGerenteGP,
    isGerenteGP: isGerenteGP,
    isCoordenador: isCoordenador
  };
}

/**
 * Valida se o usuário tem jurisdição sobre uma determinada loja por Regional
 */
function validarAcessoFilial(email, filialId) {
  var controle = obterControleAcesso(email);
  if (!controle.temAcesso) return false;
  if (controle.isSuperAdmin) return true;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var abaLojas = ss.getSheetByName('DADOS_LOJAS');
  if (!abaLojas) return false;

  var dadosLojas = abaLojas.getDataRange().getValues();
  var regionalLoja = '';

  for (var i = 1; i < dadosLojas.length; i++) {
    var idFormatado = ("0000" + dadosLojas[i][0]).slice(-4);
    var targetIdFormatado = ("0000" + filialId).slice(-4);
    if (idFormatado === targetIdFormatado) {
      regionalLoja = String(dadosLojas[i][2] || '').trim().toUpperCase();
      break;
    }
  }

  if (!regionalLoja) return false;
  return controle.regionais.includes(regionalLoja);
}
