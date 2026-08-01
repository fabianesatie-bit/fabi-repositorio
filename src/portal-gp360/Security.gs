/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Security.gs
 * Subpasta Monorepo: src/portal-gp360/
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

function obterControleAcesso(email) {
  var emailAtivo = email || Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var abaUsuarios = ss.getSheetByName('DADOS_USUARIOS');
  
  if (!abaUsuarios) {
    return { temAcesso: false, motivo: 'Aba DADOS_USUARIOS não encontrada.' };
  }
  
  var dados = abaUsuarios.getDataRange().getValues();
  var usuarioEncontrado = null;
  
  for (var i = 1; i < dados.length; i++) {
    var emailRow = String(dados[i][0] || '').toLowerCase().trim();
    if (emailRow === String(emailAtivo).toLowerCase().trim()) {
      usuarioEncontrado = {
        email: emailAtivo,
        nome: dados[i][1] || 'Usuário GP',
        cargo: dados[i][2] || 'Colaborador',
        diretoria: dados[i][3] || 'Geral',
        regionaisStr: String(dados[i][4] || ''),
        nivelAcesso: String(dados[i][5] || '').trim(),
        fotoUrl: dados[i][6] || '',
        observacoesAcesso: dados[i][7] || ''
      };
      break;
    }
  }

  if (!usuarioEncontrado) {
    registrarAuditoria('BLOQUEIO_ACESSO', 'E-mail ' + emailAtivo + ' não cadastrado.');
    return { temAcesso: false, email: emailAtivo, motivo: 'Usuário não cadastrado na base oficial do Portal GP 360.' };
  }

  var nivelNorm = usuarioEncontrado.nivelAcesso.toLowerCase();
  var ehPermitido = PERMITTED_ROLES.some(function(role) {
    return nivelNorm.includes(role);
  });

  if (!ehPermitido) {
    registrarAuditoria('BLOQUEIO_PERFIL', 'Acesso negado para ' + emailAtivo + ' com perfil: ' + usuarioEncontrado.nivelAcesso);
    return { 
      temAcesso: false, 
      email: emailAtivo, 
      motivo: 'Acesso restrito. Seu perfil (' + usuarioEncontrado.nivelAcesso + ') não possui permissão para este portal.' 
    };
  }

  var isSuperAdmin = (nivelNorm.includes('administrador'));
  var isGerenteGP = (nivelNorm.includes('gerenterh'));
  var isCoordenador = (nivelNorm.includes('coordenador'));

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

function validarAcessoFilial(email, filialId) {
  var controle = obterControleAcesso(email);
  if (!controle.temAcesso) return false;
  if (controle.isSuperAdmin) return true;

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var abaLojas = ss.getSheetByName('DADOS_LOJAS');
  if (!abaLojas) return false;

  var dadosLojas = abaLojas.getDataRange().getValues();
  var regionalLoja = '';
  var targetNorm = normalizarFilialId(filialId);

  for (var i = 1; i < dadosLojas.length; i++) {
    var norm = normalizarFilialId(dadosLojas[i][0]);
    if (norm === targetNorm) {
      regionalLoja = String(dadosLojas[i][2] || '').trim().toUpperCase();
      break;
    }
  }

  if (!regionalLoja) return false;
  return controle.regionais.includes(regionalLoja);
}
