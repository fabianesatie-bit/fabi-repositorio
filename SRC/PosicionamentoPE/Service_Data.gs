// =====================================================================
// SERVIÇO DE DADOS E VALIDAÇÃO DE IDENTIDADE
// =====================================================================

// Regra Global Obrigatória GP360: Normalização de Filiais
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

function obterInfoSessao() {
  var email = Session.getActiveUser().getEmail();
  if (!email) email = "E-mail não identificado (Execute como usuário acessando)";
  return { emailLogado: email };
}

function validarEBuscarID(idCracha) {
  try {
    var idStr = String(idCracha).trim();
    
    // --- NOVO: DOUBLE-LOCK (Trava Frontal de UX para Duplicidade) ---
    // Checa a base de respostas ANTES de abrir o formulário
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    
    if (sheetDestino) {
      var dadosDestino = sheetDestino.getDataRange().getValues();
      // O ID fica na Coluna L (índice 11)
      for (var d = 1; d < dadosDestino.length; d++) {
        if (String(dadosDestino[d][11]).trim() === idStr) {
          return { 
            sucesso: false, 
            msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail para: " + EMAIL_AGENCIA 
          };
        }
      }
    }
    // ----------------------------------------------------------------

    var emailLogado = Session.getActiveUser().getEmail();
    var isSuperAdmin = SUPER_ADMINS.indexOf(emailLogado) !== -1;
    
    var ssBase = SpreadsheetApp.openById(SPREADSHEET_BASE_ID);
    var sheetAtivos = ssBase.getSheetByName('DADOS_ATIVOS');
    
    if (!sheetAtivos) return { sucesso: false, msg: "Erro: Aba DADOS_ATIVOS não encontrada na base." };
    
    // Leitura em lote para não estourar cota (Pilar Volumetria)
    var dados = sheetAtivos.getDataRange().getValues();
    var colaboradorEncontrado = null;
    
    for (var i = 1; i < dados.length; i++) {
      var idLinha = String(dados[i][1]).trim();
      if (idLinha === idStr) {
        colaboradorEncontrado = {
          nome: dados[i][2],
          filial: normalizarFilialId(dados[i][3]),
          regional: dados[i][5],
          situacao: dados[i][6],
          cargo: dados[i][7], // Coluna H
          cpf: dados[i][11],  // Coluna L
          rg: dados[i][12],   // Coluna M
          orgao: dados[i][13], // Coluna N
          nascimento: dados[i][14] // Coluna O
        };
        break;
      }
    }
    
    if (!colaboradorEncontrado) {
      return { sucesso: false, msg: "Matrícula não encontrada na base de ativos." };
    }

    if (!isSuperAdmin) {
       var situacao = String(colaboradorEncontrado.situacao).toUpperCase();
       if (situacao.indexOf("AFASTAMENTO") !== -1 || situacao.indexOf("AUXÍLIO") !== -1 || situacao.indexOf("INSS") !== -1) {
          return { sucesso: false, msg: "Acesso bloqueado: Colaborador em afastamento." };
       }

       // A Regional "ESCRITORIO" foi mantida na regra de validação
       var regionaisPermitidas = ["BRASILIA", "UBERLANDIA", "JUIZ DE FORA", "BELO HORIZONTE", "CAMPO GRANDE", "CUIABA", "ESCRITORIO"];
       var regionalColab = String(colaboradorEncontrado.regional).toUpperCase().replace(/[ÁÀÂÃ]/g, 'A').replace(/[ÍÌÎ]/g, 'I');
       
       var pertenceAoPiloto = false;
       for(var r=0; r<regionaisPermitidas.length; r++){
           if(regionalColab.indexOf(regionaisPermitidas[r]) !== -1){
               pertenceAoPiloto = true; break;
           }
       }
       
       if (!pertenceAoPiloto) {
         return { sucesso: false, msg: "Sua regional (" + colaboradorEncontrado.regional + ") não faz parte deste piloto." };
       }
    }

    return { 
      sucesso: true, 
      dados: {
         nome: colaboradorEncontrado.nome,
         regional: colaboradorEncontrado.regional,
         filial: colaboradorEncontrado.filial,
         cargo: colaboradorEncontrado.cargo,
         cpf: colaboradorEncontrado.cpf,
         rg: colaboradorEncontrado.rg,
         orgao: colaboradorEncontrado.orgao,
         nascimento: colaboradorEncontrado.nascimento
      }
    };

  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor ao buscar dados: " + e.message };
  }
}
