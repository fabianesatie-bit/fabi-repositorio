// =====================================================================
// SERVIÇO DE DADOS E VALIDAÇÃO DE IDENTIDADE - ECOSSISTEMA GP360
// =====================================================================

// Regra Global Obrigatória GP360: Normalização de Filiais (Middleware Local)
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

// Vacina contra falha de serialização (Evita erro de Date object retornando pro HTML)
function formatarDataSegura(valor) {
  if (!valor) return "";
  if (valor instanceof Date) {
    try {
      return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } catch(e) {
      return String(valor);
    }
  }
  return String(valor);
}

function obterInfoSessao() {
  var email = Session.getActiveUser().getEmail();
  if (!email) email = "E-mail não identificado (Execute como usuário acessando)";
  return { emailLogado: email };
}

/**
 * Função de roteamento e busca com Double-Lock e seleção de Origem
 * @param {string} idCracha - Matrícula do colaborador
 * @param {string} origem - 'LOJAS' ou 'ESCRITORIO'
 */
function validarEBuscarID(idCracha, origem) {
  try {
    if (!idCracha || !origem) {
      return { sucesso: false, msg: "Matrícula ou Origem ausentes." };
    }

    var emailLogado = Session.getActiveUser().getEmail();
    var isSuperAdmin = SUPER_ADMINS.indexOf(emailLogado) !== -1;
    
    // 1. DOUBLE-LOCK (TRAVA FRONTAL ANTI-DUPLICIDADE)
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    if (sheetDestino) {
      var dadosDestino = sheetDestino.getDataRange().getValues();
      for (var k = 1; k < dadosDestino.length; k++) {
        var matriculaExistente = String(dadosDestino[k][11]).trim(); // Coluna L (Índice 11)
        if (matriculaExistente === String(idCracha).trim()) {
           return { 
             sucesso: false, 
             msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail para: " + EMAIL_AGENCIA 
           };
        }
      }
    }

    // 2. DIRECIONAMENTO DE BUSCA (PERFORMANCE E ISOLAMENTO)
    var ssBase = SpreadsheetApp.openById(SPREADSHEET_BASE_ID);
    var colaboradorEncontrado = null;

    if (origem === 'LOJAS') {
      var sheetAtivos = ssBase.getSheetByName('DADOS_ATIVOS');
      if (!sheetAtivos) return { sucesso: false, msg: "Aba DADOS_ATIVOS não encontrada." };
      
      var dados = sheetAtivos.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        var idLinha = String(dados[i][1]).trim();
        if (idLinha === String(idCracha).trim()) {
          colaboradorEncontrado = {
            nome: dados[i][2],
            filial: normalizarFilialId(dados[i][3]),
            regional: dados[i][5],
            situacao: dados[i][6],
            cargo: dados[i][7], // Coluna H
            cpf: dados[i][11],  // Coluna L
            rg: dados[i][12],   // Coluna M
            orgao: dados[i][13], // Coluna N
            nascimento: formatarDataSegura(dados[i][14]) // Vacina de data
          };
          break;
        }
      }
    } 
    else if (origem === 'ESCRITORIO') {
      var sheetEscritorio = ssBase.getSheetByName('DADOS_ESCRITORIO');
      if (!sheetEscritorio) return { sucesso: false, msg: "Aba DADOS_ESCRITORIO não encontrada." };
      
      var dadosEsc = sheetEscritorio.getDataRange().getValues();
      for (var j = 1; j < dadosEsc.length; j++) {
        var idLinhaEsc = String(dadosEsc[j][6]).trim(); // Coluna 7 (Índice 6) - ID
        if (idLinhaEsc === String(idCracha).trim()) {
          colaboradorEncontrado = {
            nome: dadosEsc[j][7], // Coluna 8 (Índice 7) - NOME
            filial: normalizarFilialId(dadosEsc[j][0]), // Coluna 1 (Índice 0) - FILIAL
            regional: dadosEsc[j][4], // Coluna 5 (Índice 4) - REGIONAL
            situacao: dadosEsc[j][24], // Coluna 25 (Índice 24) - SITUACAO
            cargo: dadosEsc[j][17], // Coluna 18 (Índice 17) - CARGO
            cpf: dadosEsc[j][11], // Coluna 12 (Índice 11) - CPF
            rg: dadosEsc[j][12], // Coluna 13 (Índice 12) - RG
            orgao: dadosEsc[j][13], // Coluna 14 (Índice 13) - ORGAO
            nascimento: formatarDataSegura(dadosEsc[j][23]) // Vacina de data aplicada!
          };
          break;
        }
      }
    }

    if (!colaboradorEncontrado) {
      return { sucesso: false, msg: "Matrícula não encontrada na base selecionada (" + origem + ")." };
    }

    // 3. REGRAS DE NEGÓCIO E PERMISSÕES DE ACESSO
    if (!isSuperAdmin) {
       var situacao = String(colaboradorEncontrado.situacao).toUpperCase();
       if (situacao.indexOf("AFASTAMENTO") !== -1 || situacao.indexOf("AUXÍLIO") !== -1 || situacao.indexOf("INSS") !== -1) {
          return { sucesso: false, msg: "Acesso bloqueado: Colaborador em afastamento." };
       }

       if (origem === 'LOJAS') {
           var regionaisPermitidas = ["BRASILIA", "UBERLANDIA", "JUIZ DE FORA", "BELO HORIZONTE", "CAMPO GRANDE", "CUIABA", "ESCRITORIO"];
           var regionalColab = String(colaboradorEncontrado.regional).toUpperCase().replace(/[ÁÀÂÃ]/g, 'A').replace(/[ÍÌÎ]/g, 'I');
           
           var pertenceAoPiloto = false;
           for(var r = 0; r < regionaisPermitidas.length; r++) {
               if(regionalColab.indexOf(regionaisPermitidas[r]) !== -1){
                   pertenceAoPiloto = true; break;
               }
           }
           
           if (!pertenceAoPiloto) {
             return { sucesso: false, msg: "Sua regional (" + colaboradorEncontrado.regional + ") não faz parte deste piloto." };
           }
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
