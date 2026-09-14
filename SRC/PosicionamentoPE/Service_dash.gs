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
    // 1. Ler Base de Preenchimentos (Destino)
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    if (!sheetDestino) return { sucesso: false, erro: "Aba 'Detalhamento Regional' não encontrada." };
    
    var dadosPreenchidos = sheetDestino.getDataRange().getValues();
    var idsPreenchidos = [];
    
    for (var i = 1; i < dadosPreenchidos.length; i++) {
       var id = String(dadosPreenchidos[i][11]).trim();
       if (id !== "") idsPreenchidos.push(id);
    }
    
    // 2. Ler Base de Ativos Corporativa (Origem)
    var ssBase = SpreadsheetApp.openById(SPREADSHEET_BASE_ID);
    var sheetAtivos = ssBase.getSheetByName('DADOS_ATIVOS');
    if (!sheetAtivos) return { sucesso: false, erro: "Aba 'DADOS_ATIVOS' não encontrada." };
    
    var dadosAtivos = sheetAtivos.getDataRange().getValues();
    
    // --- ADICIONADO: ESCRITORIO ---
    var regionaisPiloto = ["BRASILIA", "UBERLANDIA", "JUIZ DE FORA", "BELO HORIZONTE", "CAMPO GRANDE", "CUIABA", "ESCRITORIO"];
    var listaGerentes = [];
    var totalPreenchidos = 0;
    var totalPendentes = 0;

    for (var j = 1; j < dadosAtivos.length; j++) {
      var idAtivo = String(dadosAtivos[j][1]).trim();
      var nomeAtivo = dadosAtivos[j][2];
      var filialAtivo = normalizarFilialIdDash(dadosAtivos[j][3]);
      var regionalAtivoOrig = String(dadosAtivos[j][5]);
      var regionalNormalizada = regionalAtivoOrig.toUpperCase().replace(/[ÁÀÂÃ]/g, 'A').replace(/[ÍÌÎ]/g, 'I');
      var situacao = String(dadosAtivos[j][6]).toUpperCase();
      var cargo = String(dadosAtivos[j][7]).toUpperCase();
      var cargoResumido = String(dadosAtivos[j][8]).toUpperCase();
      var loginEmail = String(dadosAtivos[j][9]).trim() + "@magazineluiza.com.br";
      
      var isGerente = (cargo.indexOf("GERENTE") !== -1 || cargoResumido.indexOf("GERENTE") !== -1);
      var isAfastado = (situacao.indexOf("AFASTAMENTO") !== -1 || situacao.indexOf("AUXÍLIO") !== -1 || situacao.indexOf("INSS") !== -1);
      
      var pertenceAoPiloto = false;
      for (var r = 0; r < regionaisPiloto.length; r++) {
         if (regionalNormalizada.indexOf(regionaisPiloto[r]) !== -1) {
             pertenceAoPiloto = true; break;
         }
      }
      
      if (isGerente && pertenceAoPiloto && !isAfastado) {
         var statusPreenchimento = (idsPreenchidos.indexOf(idAtivo) !== -1) ? "Preenchido OK" : "Pendente";
         
         if (statusPreenchimento === "Preenchido OK") totalPreenchidos++;
         else totalPendentes++;
         
         listaGerentes.push({
            id: idAtivo,
            nome: nomeAtivo,
            regional: regionalAtivoOrig,
            filial: filialAtivo,
            status: statusPreenchimento,
            email: loginEmail
         });
      }
    }
    
    var isSuperAdmin = (SUPER_ADMINS.indexOf(Session.getActiveUser().getEmail()) !== -1);

    var payload = {
      gerentes: listaGerentes,
      totalGeral: listaGerentes.length,
      preenchidos: totalPreenchidos,
      pendentes: totalPendentes,
      isSuperAdmin: isSuperAdmin
    };

    return { sucesso: true, payload: payload };
    
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}
