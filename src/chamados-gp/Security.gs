/**
 * ============================================================================
 * MÓDULO: Security.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Security.gs)
 * ============================================================================
 */

// Lista fixa de e-mails autorizados para acesso ao Dashboard e recebimento de alertas
var EMAILS_AUTORIZADOS_COORD = [
  "lorena.castro@magazineluiza.com.br",
  "fabiane.satie@magazineluiza.com.br",
  "milena.ramos@magazineluiza.com.br",
  "marcia.almeida@magazineluiza.com.br",
  "josiane.emilio@magazineluiza.com.br",
  "wylmer.garcia@magazineluiza.com.br",
  "emerson.matos@magazineluiza.com.br",
  "anne.dinis@magazineluiza.com.br",
  "nayara.moraes@magazineluiza.com.br",
  "bianca.crespo@magazineluiza.com.br",
  "lelia.castelao@magazineluiza.com.br",
  "maria.antunes@magazineluiza.com.br",
  "luiz.crepaldi@magazineluiza.com.br",
  "lenilton.silveira@magazineluiza.com.br",
  "soraya.yoshida@magazineluiza.com.br",
  "fernanda.sampaio@magazineluiza.com.br",
  "gua_santos@magazineluiza.com.br",
  "marcos.costa@magazineluiza.com.br",
  "mario.soares@magazineluiza.com.br",
  "bruna.nayra@magazineluiza.com.br",
  "jamile.sales@magazineluiza.com.br",
  "ticyanny.figueredo@magazineluiza.com.br",
  "apoio_mg_co@magazineluiza.com.br"
];

/**
 * Valida se o usuário atual possui permissão de Coordenador ou Gerente de GP.
 * @param {string} [chaveAcesso] Parâmetro opcional passado via URL
 * @return {Object} Objeto com status da permissao, email e papel do usuario
 */
function validarAcessoCoordenador(chaveAcesso) {
  var emailAtivo = "";
  try {
    emailAtivo = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (e) {
    emailAtivo = "";
  }

  // 1. Verificação por Chave Global Autorizada
  if (chaveAcesso && (chaveAcesso.toLowerCase() === "gp" || chaveAcesso.toLowerCase() === "mgco" || chaveAcesso.toLowerCase() === "sul")) {
    return {
      autorizado: true,
      email: emailAtivo || "coordenador.autorizado@magazineluiza.com.br",
      cargo: "Coordenador/Gerente GP",
      metodo: "ChaveAcesso"
    };
  }

  // 2. Verificação por E-mail do Usuário na Lista Fixa
  if (emailAtivo !== "") {
    for (var i = 0; i < EMAILS_AUTORIZADOS_COORD.length; i++) {
      if (EMAILS_AUTORIZADOS_COORD[i].toLowerCase() === emailAtivo) {
        return {
          autorizado: true,
          email: emailAtivo,
          cargo: emailAtivo === "fabiane.satie@magazineluiza.com.br" ? "Gerente RH / Dev" : "Coordenador GP",
          metodo: "EmailSessao"
        };
      }
    }
  }

  // 3. Verificação Dinâmica na Aba DADOS_USUARIOS da Planilha Principal
  try {
    var ssMaster = SpreadsheetApp.openById(CONFIG.PLANILHA_MASTER_ID);
    var abaUsuarios = ssMaster.getSheetByName(CONFIG.ABA_USUARIOS);
    if (abaUsuarios && emailAtivo !== "") {
      var dados = abaUsuarios.getDataRange().getValues();
      for (var r = 1; r < dados.length; r++) {
        var emailBd = String(dados[r][0] || "").toLowerCase().trim();
        var cargoBd = String(dados[r][2] || "").toLowerCase().trim();
        var nivelBd = String(dados[r][5] || "").toLowerCase().trim();

        if (emailBd === emailAtivo) {
          if (cargoBd.indexOf("gerente") !== -1 || cargoBd.indexOf("coordenador") !== -1 || nivelBd.indexOf("admin") !== -1) {
            return {
              autorizado: true,
              email: emailAtivo,
              cargo: dados[r][2] || "GP",
              metodo: "PlanilhaUsuarios"
            };
          }
        }
      }
    }
  } catch (err) {
    Logger.log("Erro ao consultar DADOS_USUARIOS: " + err.toString());
  }

  return {
    autorizado: false,
    email: emailAtivo,
    cargo: "Solicitante",
    metodo: "Negado"
  };
}

/**
 * Valida se o usuário logado é a desenvolvedora do sistema.
 * @return {boolean}
 */
function validarAcessoDesenvolvedora() {
  var emailAtivo = "";
  try {
    emailAtivo = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (e) {
    emailAtivo = "";
  }
  
  if (emailAtivo === CONFIG.EMAIL_DEV.toLowerCase()) {
    return true;
  }
  
  var checagem = validarAcessoCoordenador();
  return checagem.autorizado;
}
