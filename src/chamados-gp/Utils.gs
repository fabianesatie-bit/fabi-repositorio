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
 * @param {string|number} val
 * @return {string}
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
 * Sanitiza matrículas mantendo apenas números.
 * @param {string|number} input
 * @return {string}
 */
function sanitizarMatriculaId(input) {
  if (!input && input !== 0) return "";
  return String(input).replace(/\D/g, '').trim();
}

/**
 * Converte automaticamente qualquer URL presente em um texto em link HTML clicável (target="_blank").
 * @param {string} texto
 * @return {string}
 */
function converterLinksParaHtml(texto) {
  if (!texto) return "";
  var regexUrl = /(https?:\/\/[^\s<]+)/g;
  return String(texto).replace(regexUrl, function(url) {
    return '<a href="' + url + '" target="_blank" style="color: #2563EB; font-weight: 600; text-decoration: underline;">' + url + '</a>';
  });
}

/**
 * Gera o próximo número de protocolo sequencial para o ano vigente.
 * Estrutura: GP-[ANO]-[NÚMERO COM 5 DÍGITOS] (Ex: GP-2026-00001)
 * @param {Sheet} abaBase
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
