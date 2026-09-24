// =====================================================================
// SERVIÇO DE DADOS E VALIDAÇÃO DE IDENTIDADE - ECOSSISTEMA GP360
// MÓDULO: FORMULÁRIO DE INSCRIÇÃO (BUSCA UNIFICADA SIMPLIFICADA)
// =====================================================================

// ---------------------------------------------------------------------
// REGRA GLOBAL OBRIGATÓRIA GP360: NORMALIZAÇÃO DE FILIAIS (MIDDLEWARE LOCAL)
// ---------------------------------------------------------------------
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

// ---------------------------------------------------------------------
// FUNÇÃO AUXILIAR OBRIGATÓRIA: SANITIZAÇÃO E REMOÇÃO COMPLETA DE ACENTOS
// ---------------------------------------------------------------------
function sanitizarTexto(texto) {
  if (!texto) return "";
  return String(texto)
    .toUpperCase()
    .trim()
    .replace(/[ÁÀÂÃ]/g, 'A')
    .replace(/[ÉÈÊ]/g, 'E')
    .replace(/[ÍÌÎ]/g, 'I')
    .replace(/[ÓÒÔÕ]/g, 'O')
    .replace(/[ÚÙÛ]/g, 'U')
    .replace(/Ç/g, 'C');
}

// ---------------------------------------------------------------------
// VACINA CONTRA FALHA DE SERIALIZAÇÃO (EVITA ERRO DE DATE OBJECT NO HTML)
// ---------------------------------------------------------------------
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
  if (!email) email = "E-mail não identificado";
  return { emailLogado: email };
}

/**
 * BUSCA UNIFICADA APENAS POR ID (MÓDULO SIMPLIFICADO GP360)
 * Varre primeiramente DADOS_ATIVOS; se não localizar, varre DADOS_ESCRITORIO.
 * @param {string} idCracha - Matrícula do colaborador
 */
function validarEBuscarID(idCracha) {
  try {
    if (!idCracha) {
      return { sucesso: false, msg: "Por favor, digite o ID do crachá/matrícula." };
    }

    var idCrachaLimpo = String(idCracha).trim();
    var emailLogado = Session.getActiveUser().getEmail();
    var isSuperAdmin = (typeof SUPER_ADMINS !== 'undefined' && SUPER_ADMINS.indexOf(emailLogado) !== -1);
    
    // 1. DOUBLE-LOCK (TRAVA FRONTAL ANTI-DUPLICIDADE)
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    if (sheetDestino) {
      var dadosDestino = sheetDestino.getDataRange().getValues();
      for (var k = 1; k < dadosDestino.length; k++) {
        var matriculaExistente = String(dadosDestino[k][11]).trim(); // Coluna L (Índice 11)
        if (matriculaExistente === idCrachaLimpo) {
           return { 
             sucesso: false, 
             msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail para a equipe de organização."
           };
        }
      }
    }

    var ssBase = SpreadsheetApp.openById(SPREADSHEET_BASE_ID);
    var colaboradorEncontrado = null;
    var tipoOrigem = "";

    // 2. PASSO A: BUSCA NA ABA DADOS_ATIVOS (LOJAS)
    var sheetAtivos = ssBase.getSheetByName('DADOS_ATIVOS');
    if (sheetAtivos) {
      var mapaLojas = {};
      var sheetLojas = ssBase.getSheetByName('DADOS_LOJAS');
      if (sheetLojas) {
        var dadosLojas = sheetLojas.getDataRange().getValues();
        for (var l = 1; l < dadosLojas.length; l++) {
          var idFilialNorm = normalizarFilialId(dadosLojas[l][0]); // Coluna A (Filial_ID)
          if (idFilialNorm !== "") {
            mapaLojas[idFilialNorm] = String(dadosLojas[l][2] || "").trim(); // Coluna C (Regional Real)
          }
        }
      }

      var dadosAtivos = sheetAtivos.getDataRange().getValues();
      for (var i = 1; i < dadosAtivos.length; i++) {
        var idLinha = String(dadosAtivos[i][1]).trim(); // Coluna B (Id)
        if (idLinha === idCrachaLimpo) {
          var filialNorm = normalizarFilialId(dadosAtivos[i][3]); // Coluna D (filial)
          var regionalOficial = mapaLojas[filialNorm] || String(dadosAtivos[i][5] || "").trim();

          colaboradorEncontrado = {
            id: idLinha,
            nome: dadosAtivos[i][2],
            filial: filialNorm,
            regional: regionalOficial,
            situacao: dadosAtivos[i][6],
            cargo: dadosAtivos[i][7], // Coluna H
            cpf: dadosAtivos[i][11],  // Coluna L
            rg: dadosAtivos[i][12],   // Coluna M
            orgao: dadosAtivos[i][13], // Coluna N
            nascimento: formatarDataSegura(dadosAtivos[i][14])
          };
          tipoOrigem = "LOJAS";
          break;
        }
      }
    }

    // 3. PASSO B: FALLBACK AUTOMÁTICO NA ABA DADOS_ESCRITORIO (CORPORATIVO / REGIONAL)
    if (!colaboradorEncontrado) {
      var sheetEscritorio = ssBase.getSheetByName('DADOS_ESCRITORIO');
      if (sheetEscritorio) {
        var dadosEsc = sheetEscritorio.getDataRange().getValues();
        for (var j = 1; j < dadosEsc.length; j++) {
          var idLinhaEsc = String(dadosEsc[j][6]).trim(); // Coluna G (Index 6) - CHAPA/ID
          if (idLinhaEsc === idCrachaLimpo) {
            colaboradorEncontrado = {
              id: idLinhaEsc,
              nome: dadosEsc[j][7], // Coluna H (Index 7) - NOME
              filial: normalizarFilialId(dadosEsc[j][0]), // Coluna A (Index 0) - FILIAL
              regional: String(dadosEsc[j][4] || "").trim(), // Coluna E (Index 4) - REGIONAL
              situacao: dadosEsc[j][24], // Coluna Y (Index 24) - SITUACAO
              cargo: dadosEsc[j][17], // Coluna R (Index 17) - CARGO
              cpf: dadosEsc[j][11], // Coluna L (Index 11) - CPF
              rg: dadosEsc[j][12], // Coluna M (Index 12) - RG
              orgao: dadosEsc[j][13], // Coluna N (Index 13) - ORGAO
              nascimento: formatarDataSegura(dadosEsc[j][23])
            };
            tipoOrigem = "ESCRITORIO";
            break;
          }
        }
      }
    }

    // SE NÃO ENCONTRADO EM NENHUMA DAS DUAS ABAS
    if (!colaboradorEncontrado) {
      return { 
        sucesso: false, 
        msg: "Matrícula " + idCrachaLimpo + " não localizada na base. Verifique o número digitado." 
      };
    }

    // 4. VERIFICAÇÃO DE AFASTAMENTO DO RH
    if (!isSuperAdmin) {
       var situacaoSanitizada = sanitizarTexto(colaboradorEncontrado.situacao);
       if (situacaoSanitizada.indexOf("AFASTAMENTO") !== -1 || situacaoSanitizada.indexOf("AUXILIO") !== -1 || situacaoSanitizada.indexOf("INSS") !== -1) {
          return { sucesso: false, msg: "Acesso bloqueado: Colaborador em afastamento no RH." };
       }
    }

    return { 
      sucesso: true, 
      dados: {
         id: colaboradorEncontrado.id,
         nome: colaboradorEncontrado.nome,
         regional: colaboradorEncontrado.regional,
         filial: colaboradorEncontrado.filial,
         cargo: colaboradorEncontrado.cargo,
         cpf: colaboradorEncontrado.cpf,
         rg: colaboradorEncontrado.rg,
         orgao: colaboradorEncontrado.orgao,
         nascimento: colaboradorEncontrado.nascimento,
         origem: tipoOrigem,
         requerRegionalManual: (tipoOrigem === "ESCRITORIO")
      }
    };

  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor ao buscar dados: " + e.message };
  }
}
