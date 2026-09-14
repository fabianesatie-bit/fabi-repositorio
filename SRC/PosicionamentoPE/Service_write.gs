// =====================================================================
// SERVIÇO DE GRAVAÇÃO E VALIDAÇÃO ANTI-DUPLICIDADE
// =====================================================================

// Regra Global GP360: Normalização de Filiais (Middleware Local)
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

function salvarDados(payload) {
  try {
    // BLINDAGEM DE INTEGRIDADE (BACK-END): Se o transporte não for aéreo, zera aeroportos à força.
    if (payload.tipoTransporte !== 'Aéreo') {
      payload.aeroporto1 = "";
      payload.aeroporto2 = "";
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheet = ss.getSheetByName('Detalhamento Regional');
    
    if (!sheet) {
      return { sucesso: false, msg: "ERRO CRÍTICO: Aba 'Detalhamento Regional' não encontrada." };
    }
    
    var dados = sheet.getDataRange().getValues();
    
    // Verificação Anti-Duplicidade (ID fica na Coluna L - índice 11)
    for (var i = 1; i < dados.length; i++) {
      var matriculaExistente = String(dados[i][11]).trim(); 
      if (matriculaExistente === String(payload.id).trim()) {
        return { 
          sucesso: false, 
          msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail diretamente para: " + EMAIL_AGENCIA 
        };
      }
    }
    
    // Preparação para Gravação - A e B vazias para a Agência
    var novaLinha = new Array(28).fill("");
    
    novaLinha[0] = ""; // Coluna A - VAZIA (Agência)
    novaLinha[1] = ""; // Coluna B - VAZIA (Agência)
    novaLinha[2] = payload.regional; // Coluna C - REGIONAL
    novaLinha[3] = normalizarFilialId(payload.filial); // Coluna D - FILIAL
    novaLinha[4] = payload.nomeCompleto; // Coluna E - NOME
    novaLinha[5] = payload.sexo; // Coluna F - SEXO
    novaLinha[6] = payload.rg; // Coluna G - RG
    novaLinha[7] = payload.orgao; // Coluna H - ÓRGÃO
    novaLinha[8] = payload.cpf; // Coluna I - CPF
    
    // Formatação segura de Data
    var dataNasc = payload.nascimento;
    if (dataNasc && typeof dataNasc.getTime === 'function') {
      novaLinha[9] = Utilities.formatDate(dataNasc, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      novaLinha[9] = dataNasc; // Coluna J - NASCIMENTO
    }
    
    novaLinha[10] = payload.cargo; // Coluna K - CARGO
    novaLinha[11] = payload.id; // Coluna L - ID CRACHÁ
    novaLinha[12] = payload.ddd; // Coluna M - DDD
    novaLinha[13] = payload.celular; // Coluna N - CELULAR
    novaLinha[14] = ""; // Coluna O - HOSPEDAGEM (Oculto no momento)
    novaLinha[15] = ""; // Coluna P - DATA HOSPEDAGEM (Oculto)
    novaLinha[16] = ""; // Coluna Q - DUPLA DE QUARTO (Oculto)
    novaLinha[17] = payload.restricaoAlimentar; // Coluna R - RESTRIÇÃO
    novaLinha[18] = payload.emailMagalu; // Coluna S - EMAIL CORP
    novaLinha[19] = payload.emailPessoal; // Coluna T - EMAIL PESSOAL
    novaLinha[20] = payload.tipoTransporte; // Coluna U - TRANSPORTE
    novaLinha[21] = payload.observacaoRotas; // Coluna V - OBS ROTAS
    novaLinha[22] = payload.aeroporto1; // Coluna W - AERO 1
    novaLinha[23] = payload.aeroporto2; // Coluna X - AERO 2
    novaLinha[24] = payload.cidadeEmbarque; // Coluna Y - CIDADE EMBARQUE
    novaLinha[25] = payload.mobilidadeReduzida; // Coluna Z - MOBILIDADE
    novaLinha[26] = payload.observacoesGerais; // Coluna AA - OBS GERAIS
    
    sheet.appendRow(novaLinha);
    
    return { sucesso: true, msg: "Dados validados e salvos com sucesso!" };
    
  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor ao salvar os dados: " + e.message };
  }
}
