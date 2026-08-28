/**
 * MOTOR DE SEGURANÇA E CONTROLE DE ACESSO (RLS)
 */

function obterPerfilAcessoLogado() {
  var emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  
  var perfil = {
    email: emailLogado,
    nome: emailLogado.split('@')[0],
    isSuperAdmin: SUPER_ADMIN_EMAILS.includes(emailLogado),
    regionaisPermitidas: [],
    diretoriasPermitidas: [],
    temAcessoTotal: false
  };

  if (perfil.isSuperAdmin) {
    perfil.temAcessoTotal = true;
    return perfil;
  }

  try {
    var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
    var shUsuarios = ssMaster.getSheetByName(TAB_NAMES.USUARIOS);
    if (shUsuarios) {
      var dados = shUsuarios.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        var emailLinha = dados[i][0] ? String(dados[i][0]).toLowerCase().trim() : '';
        if (emailLinha === emailLogado) {
          perfil.nome = dados[i][1] || perfil.nome;
          var cargo = String(dados[i][2] || '').toUpperCase();
          var regStr = String(dados[i][4] || '').toUpperCase();
          var dirStr = String(dados[i][3] || '').toUpperCase();

          if (cargo.includes('DIRETOR') || cargo.includes('GERENTEGP')) {
            perfil.temAcessoTotal = true;
          }

          if (regStr) {
            regStr.split(',').forEach(function(r) {
              var rLimpa = r.trim();
              if (rLimpa) perfil.regionaisPermitidas.push(rLimpa);
            });
          }

          if (dirStr) {
            dirStr.split(',').forEach(function(d) {
              var dLimpo = d.trim();
              if (dLimpo) perfil.diretoriasPermitidas.push(dLimpo);
            });
          }
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Erro ao carregar segurança: " + e.message);
  }

  if (perfil.regionaisPermitidas.includes('TODAS') || perfil.regionaisPermitidas.length === 0) {
    perfil.temAcessoTotal = true;
  }

  return perfil;
}
