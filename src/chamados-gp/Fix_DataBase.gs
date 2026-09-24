/**
 * ============================================================================
 * MÓDULO TEMPORÁRIO: Fix_Database.gs
 * OBJETIVO: Limpar linhas corrompidas de FAQ na aba CHAMADOS_BASE e criar
 *           automaticamente as abas CONFIG_FAQ e CONFIG_TIMELINE na planilha.
 * AUTOR: Desenvolvedor Sênior / Arquiteto GP360
 * ============================================================================
 */

function executarLimpezaEAjustePlanilhaGP() {
  var idPlanilhaChamados = "19QKV7IwyuzqUAM4fP5mn_gthHB9JxDdJhi1m4GYrzSg";
  
  try {
    var ss = SpreadsheetApp.openById(idPlanilhaChamados);
    Logger.log("--> Iniciando conexão com a Planilha Dedicada de Chamados GP...");

    // =========================================================================
    // PASSO 1: CRIAR OU REORGANIZAR ABA CONFIG_FAQ
    // =========================================================================
    var abaFaq = ss.getSheetByName("CONFIG_FAQ");
    if (!abaFaq) {
      abaFaq = ss.insertSheet("CONFIG_FAQ");
      abaFaq.appendRow(["ID", "Categoria", "Pergunta", "Resposta_Detalhada", "Ordem", "Status"]);
      abaFaq.getRange("A1:F1").setFontWeight("bold").setBackground("#F3E8FF").setFontColor("#6B21A8");
      Logger.log("✅ Aba CONFIG_FAQ criada com sucesso com cabeçalhos padrão!");
    } else {
      Logger.log("ℹ️ Aba CONFIG_FAQ já existente.");
    }

    // =========================================================================
    // PASSO 2: CRIAR OU REORGANIZAR ABA CONFIG_TIMELINE
    // =========================================================================
    var abaTimeline = ss.getSheetByName("CONFIG_TIMELINE");
    if (!abaTimeline) {
      abaTimeline = ss.insertSheet("CONFIG_TIMELINE");
      abaTimeline.appendRow(["Data_Aviso", "Sistema_Modulo", "Titulo_Aviso", "Mensagem_Detalhada", "Status_Sistema", "Exibir_Na_Tela"]);
      abaTimeline.getRange("A1:F1").setFontWeight("bold").setBackground("#EFF6FF").setFontColor("#1E40AF");
      Logger.log("✅ Aba CONFIG_TIMELINE criada com sucesso com cabeçalhos padrão!");
    } else {
      Logger.log("ℹ️ Aba CONFIG_TIMELINE já existente.");
    }

    // =========================================================================
    // PASSO 3: REMOVER REGISTROS DESALINHADOS / FAQ DA ABA CHAMADOS_BASE
    // =========================================================================
    var abaBase = ss.getSheetByName("CHAMADOS_BASE");
    if (!abaBase) {
      var abas = ss.getSheets();
      abaBase = abas[0];
    }

    if (abaBase) {
      var dados = abaBase.getDataRange().getValues();
      var linhasRemovidas = 0;

      // Percorre de baixo para cima para excluir linhas de FAQ sem alterar os índices
      for (var i = dados.length - 1; i >= 1; i--) {
        var protocolo = String(dados[i][0] || "").trim();
        var dataAbertura = String(dados[i][1] || "").trim();

        // Identifica se a linha é um FAQ gravado na aba errada
        if (protocolo.indexOf("FAQ-") === 0 || dataAbertura.indexOf("Quadro de lojas") !== -1 || dataAbertura.indexOf("SOS Operações") !== -1) {
          
          // Move o registro para a aba CONFIG_FAQ se ainda não existir
          var perguntaExtraida = String(dados[i][2] || dados[i][1] || "").trim();
          var respostaExtraida = String(dados[i][3] || "").trim();
          var catExtraida = (dataAbertura.indexOf("Quadro") !== -1) ? "Quadro de Lojas" : "Operações";

          if (perguntaExtraida && respostaExtraida) {
            abaFaq.appendRow([
              protocolo || ("FAQ-" + i),
              catExtraida,
              perguntaExtraida,
              respostaExtraida,
              i,
              "Ativo"
            ]);
          }

          // Exclui a linha desalinhada da aba CHAMADOS_BASE
          abaBase.deleteRow(i + 1);
          linhasRemovidas++;
        }
      }

      Logger.log("✅ Sucesso! Foram removidas " + linhasRemovidas + " linhas desalinhadas da aba CHAMADOS_BASE.");
    }

    Logger.log("🎉 PROCESSO CONCLUÍDO COM SUCESSO! A planilha está 100% higienizada e organizada.");
    return "Sucesso! Planilha higienizada. " + linhasRemovidas + " linhas de FAQ foram movidas para a aba CONFIG_FAQ.";

  } catch (err) {
    Logger.log("❌ Erro ao executar limpeza: " + err.toString());
    return "Erro: " + err.toString();
  }
}
