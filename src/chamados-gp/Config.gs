/**
 * ============================================================================
 * MÓDULO: Config.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Config.gs)
 * ============================================================================
 */

var CONFIG = {
  // 1. Planilha Dedicada para Gravação e Leitura de Chamados GP (MG/CO e SUL)
  PLANILHA_CHAMADOS_ID: "19QKV7IwyuzqUAM4fP5mn_gthHB9JxDdJhi1m4GYrzSg",
  
  // 2. Planilha Master Central (Consulta de DADOS_ATIVOS e DADOS_USUARIOS)
  PLANILHA_MASTER_ID: "1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA",
  
  // Nomes das Abas de Banco de Dados
  ABA_CHAMADOS_BASE: "CHAMADOS_BASE",
  ABA_CHAMADOS_HISTORICO: "CHAMADOS_HISTORICO",
  ABA_CATEGORIAS: "CONFIG_CATEGORIAS",
  ABA_USUARIOS: "DADOS_USUARIOS",
  ABA_ATIVOS: "DADOS_ATIVOS",
  
  // Desenvolvedora Sênior / Responsável
  EMAIL_DEV: "fabiane.satie@magazineluiza.com.br",
  
  // Categorias Padrão de Inicialização
  CATEGORIAS_PADRAO: [
    "Afastamento/Retorno",
    "Requisição de vaga",
    "Desligamento",
    "Admissão",
    "Quadro de loja",
    "Atestado medico",
    "Outros"
  ]
};

/**
 * Obtém o objeto da planilha DEDICADA de chamados GP.
 * @return {Spreadsheet}
 */
function getSpreadsheetChamados() {
  try {
    return SpreadsheetApp.openById(CONFIG.PLANILHA_CHAMADOS_ID);
  } catch (e) {
    Logger.log("Erro ao abrir planilha de chamados por ID dedicado: " + e.toString());
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Obtém o objeto da planilha MASTER CENTRAL (para consultar DADOS_ATIVOS e Usuários).
 * @return {Spreadsheet}
 */
function getSpreadsheetMaster() {
  try {
    return SpreadsheetApp.openById(CONFIG.PLANILHA_MASTER_ID);
  } catch (e) {
    Logger.log("Erro ao abrir planilha master por ID central: " + e.toString());
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Procura uma aba na planilha por nome exato, por nome insensível a maiúsculas/minúsculas ou pega a primeira aba por fallback.
 * @param {Spreadsheet} ss
 * @param {string} nomeAbaAlvo
 * @return {Sheet}
 */
function obterAbaFlexivel(ss, nomeAbaAlvo) {
  if (!ss) return null;
  
  var aba = ss.getSheetByName(nomeAbaAlvo);
  if (aba) return aba;

  var abas = ss.getSheets();
  var nomeLimpoAlvo = String(nomeAbaAlvo).trim().toLowerCase();
  
  for (var i = 0; i < abas.length; i++) {
    var nomeAbaAtual = abas[i].getName().trim().toLowerCase();
    if (nomeAbaAtual === nomeLimpoAlvo) {
      return abas[i];
    }
  }

  if (abas.length > 0) {
    Logger.log("Aba " + nomeAbaAlvo + " não encontrada diretamente. Utilizando fallback: " + abas[0].getName());
    return abas[0];
  }

  return null;
}
