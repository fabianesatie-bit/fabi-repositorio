// =====================================================================
// SERVIÇO DE PROCESSAMENTO DO DASHBOARD E CRUZAMENTO DE BASES
// =====================================================================

function normalizarFilialIdDash(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

function carregarDadosDashboard() {
  try {
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    if (!sheetDestino) return { sucesso: false, erro: "Aba 'Detalhamento Regional' não encontrada." };
    
    var dadosPreenchidos = sheetDestino.getDataRange().getValues();
    var idsPreenchidos = [];
    
    for (var i = 1; i < dadosPreenchidos.length; i++) {
       var id = String(dadosPreenchidos[i][11]).trim();
       if (id !== "") idsPreenchidos.push(id);
    }
    
    var ssBase = SpreadsheetApp.openById(SPREADSHEET_BASE_ID);
    var regionaisPiloto = ["BRASILIA", "UBERLANDIA", "JUIZ DE FORA", "BELO HORIZONTE", "CAMPO GRANDE", "CUIABA", "ESCRITORIO"];
    var listaGerentes = [];
    var totalPreenchidos = 0;
    var totalPendentes = 0;

    // 1. Processa DADOS_ATIVOS
    var sheetAtivos = ssBase.getSheetByName('DADOS_ATIVOS');
    if (sheetAtivos) {
      var dadosAtivos = sheetAtivos.getDataRange().getValues();
      for (var j = 1; j < dadosAtivos.length; j++) {
        var idAtivo = String(dadosAtivos[j][1]).trim();
        var cargo = String(dadosAtivos[j][7]).toUpperCase();
        var situacao = String(dadosAtivos[j][6]).toUpperCase();
        var regionalNormalizada = String(dadosAtivos[j][5]).toUpperCase().replace(/[ÁÀÂÃ]/g, 'A').replace(/[ÍÌÎ]/g, 'I');
        
        var isGerente = (cargo.indexOf("GERENTE") !== -1 || String(dadosAtivos[j][8]).toUpperCase().indexOf("GERENTE") !== -1);
        var isAfastado = (situacao.indexOf("AFASTAMENTO") !== -1 || situacao.indexOf("AUXÍLIO") !== -1 || situacao.indexOf("INSS") !== -1);
        
        var pertenceAoPiloto = false;
        for (var r = 0; r < regionaisPiloto.length; r++) {
           if (regionalNormalizada.indexOf(regionaisPiloto[r]) !== -1) {
               pertenceAoPiloto = true; break;
           }
        }
        
        if (isGerente && pertenceAoPiloto && !isAfastado) {
           var status = (idsPreenchidos.indexOf(idAtivo) !== -1) ? "Preenchido OK" : "Pendente";
           if (status === "Preenchido OK") totalPreenchidos++; else totalPendentes++;
           
           listaGerentes.push({
              id: idAtivo,
              nome: dadosAtivos[j][2],
              regional: dadosAtivos[j][5],
              filial: normalizarFilialIdDash(dadosAtivos[j][3]),
              status: status,
              email: String(dadosAtivos[j][9]).trim() + "@magazineluiza.com.br"
           });
        }
      }
    }

    // 2. Processa DADOS_ESCRITORIO (Empilha no mesmo array)
    var sheetEscritorio = ssBase.getSheetByName('DADOS_ESCRITORIO');
    if (sheetEscritorio) {
       var dadosEscritorio = sheetEscritorio.getDataRange().getValues();
       for (var k = 1; k < dadosEscritorio.length; k++) {
          var idEscritorio = String(dadosEscritorio[k][6]).trim(); // ID Index 6
          if (!idEscritorio) continue;
          
          var statusEscritorio = (idsPreenchidos.indexOf(idEscritorio) !== -1) ? "Preenchido OK" : "Pendente";
          if (statusEscritorio === "Preenchido OK") totalPreenchidos++; else totalPendentes++;
          
          listaGerentes.push({
             id: idEscritorio,
             nome: dadosEscritorio[k][7],
             regional: "ESCRITORIO",
             filial: normalizarFilialIdDash(dadosEscritorio[k][0]),
             status: statusEscritorio,
             email: String(dadosEscritorio[k][10]).trim() + "@magazineluiza.com.br" // LOGIN Index 10
          });
       }
    }
    
    var payload = {
      gerentes: listaGerentes,
      totalGeral: listaGerentes.length,
      preenchidos: totalPreenchidos,
      pendentes: totalPendentes,
      isSuperAdmin: (SUPER_ADMINS.indexOf(Session.getActiveUser().getEmail()) !== -1)
    };

    return { sucesso: true, payload: payload };
    
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}
