/**
 * ============================================================================
 * MÓDULO: Utils.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Utils.gs)
 * ============================================================================
 */

/**
 * REGRA GLOBAL DE NORMALIZAÇÃO DE FILIAIS (MIDDLEWARE > 3000)
 * @param {string|number} val ID ou código bruto da filial
 * @return {string} ID da filial normalizado em formato string
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { 
    num -= 3000; 
  }
  return String(num);
}

/**
 * Sanitiza e limpa matrículas ou IDs de colaboradores, mantendo apenas dígitos.
 * @param {string|number} input
 * @return {string}
 */
function sanitizarMatriculaId(input) {
  if (!input && input !== 0) return "";
  return String(input).replace(/\D/g, '').trim();
}

/**
 * Formata um objeto Date para a string padrão corporativa dd/MM/yyyy HH:mm:ss no fuso GMT-03:00.
 * @param {Date|string} dataObj
 * @param {boolean} [incluirSegundos=false]
 * @return {string}
 */
function formatarDataHoraBr(dataObj, incluirSegundos) {
  if (!dataObj) return "";
  try {
    var d = (dataObj instanceof Date) ? dataObj : new Date(dataObj);
    if (isNaN(d.getTime())) return String(dataObj);
    var mascara = incluirSegundos ? "dd/MM/yyyy HH:mm:ss" : "dd/MM/yyyy HH:mm";
    return Utilities.formatDate(d, "GMT-03:00", mascara);
  } catch (e) {
    Logger.log("Erro ao formatar data: " + e.toString());
    return String(dataObj);
  }
}

/**
 * Gera o próximo número de protocolo sequencial para o ano vigente.
 * Estrutura: GP-[ANO]-[NÚMERO COM 5 DÍGITOS] (Ex: GP-2026-00001)
 * @param {Sheet} abaBase Objeto da aba CHAMADOS_BASE
 * @return {string}
 */
function gerarProximoProtocolo(abaBase) {
  var anoAtual = new Date().getFullYear();
  var prefixo = "GP-" + anoAtual + "-";
  var ultimoNum = 0;

  if (abaBase) {
    try {
      var dados = abaBase.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        var prot = String(dados[i][0] || "").trim();
        if (prot.indexOf(prefixo) === 0) {
          var numPart = parseInt(prot.replace(prefixo, ""), 10);
          if (!isNaN(numPart) && numPart > ultimoNum) {
            ultimoNum = numPart;
          }
        }
      }
    } catch (e) {
      Logger.log("Erro ao calcular próximo protocolo: " + e.toString());
    }
  }

  var proximoNum = ultimoNum + 1;
  var strNum = ("00000" + proximoNum).slice(-5);
  return prefixo + strNum;
}
