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

// Converte datas do formato HTML (YYYY-MM-DD) para Brasileiro (DD/MM/YYYY)
function formatarDataInclusao(dataStr) {
  if (!dataStr) return "";
  if (String(dataStr).indexOf('-') === 4) {
    var p = dataStr.split('-');
    return p[2] + '/' + p[1] + '/' + p[0];
  }
  return dataStr;
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
    
    for (var i = 1; i < dados.length; i++) {
      var matriculaExistente = String(dados[i][11]).trim(); 
      if (matriculaExistente !== "" && matriculaExistente === String(payload.id).trim()) {
        return { 
          sucesso: false, 
          msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail diretamente para: " + EMAIL_AGENCIA 
        };
      }
    }
    
    var novaLinha = new Array(28).fill("");
    
    novaLinha[0] = ""; // Coluna A - VAZIA
    novaLinha[1] = ""; // Coluna B - VAZIA
    novaLinha[2] = payload.regional; 
    novaLinha[3] = normalizarFilialId(payload.filial); 
    novaLinha[4] = payload.nomeCompleto; 
    novaLinha[5] = payload.sexo; 
    novaLinha[6] = payload.rg; 
    novaLinha[7] = payload.orgao; 
    novaLinha[8] = payload.cpf; 
    
    var dataNasc = payload.nascimento;
    if (dataNasc && typeof dataNasc.getTime === 'function') {
      novaLinha[9] = Utilities.formatDate(dataNasc, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      novaLinha[9] = dataNasc;
    }
    
    novaLinha[10] = payload.cargo; 
    novaLinha[11] = payload.id; // O ID garante que o painel conte apenas a gerente!
    novaLinha[12] = payload.ddd; 
    novaLinha[13] = payload.celular; 
    novaLinha[14] = ""; 
    novaLinha[15] = ""; 
    novaLinha[16] = ""; 
    novaLinha[17] = payload.restricaoAlimentar; 
    novaLinha[18] = payload.emailMagalu; 
    novaLinha[19] = payload.emailPessoal; 
    novaLinha[20] = payload.tipoTransporte; 
    novaLinha[21] = payload.observacaoRotas; 
    novaLinha[22] = payload.aeroporto1; 
    novaLinha[23] = payload.aeroporto2; 
    novaLinha[24] = payload.cidadeEmbarque; 
    novaLinha[25] = payload.mobilidadeReduzida; 
    novaLinha[26] = payload.observacoesGerais; 
    
    sheet.appendRow(novaLinha);
    
    if (payload.isLactante === 'Sim') {
      // Cria a linha 2 (Acompanhante)
      var linhaAcomp = new Array(28).fill("");
      linhaAcomp[2] = payload.regional;
      linhaAcomp[3] = normalizarFilialId(payload.filial);
      linhaAcomp[4] = payload.nomeAcompanhante;
      linhaAcomp[5] = "Não Informado";
      linhaAcomp[6] = payload.rgAcompanhante;
      linhaAcomp[8] = payload.cpfAcompanhante;
      linhaAcomp[9] = formatarDataInclusao(payload.nascAcompanhante);
      linhaAcomp[10] = "ACOMPANHANTE LACTANTE"; 
      linhaAcomp[11] = ""; // ID VAZIO (Ignora no Dashboard)
      linhaAcomp[12] = payload.ddd; // Usa o contato da gerente
      linhaAcomp[13] = payload.celular;
      linhaAcomp[18] = payload.emailMagalu; // Amarração de dados
      linhaAcomp[20] = payload.tipoTransporte;
      linhaAcomp[22] = payload.aeroporto1;
      linhaAcomp[23] = payload.aeroporto2;
      linhaAcomp[24] = payload.cidadeEmbarque;
      
      sheet.appendRow(linhaAcomp);

      // Cria a linha 3 (Bebê)
      var linhaBebe = new Array(28).fill("");
      linhaBebe[2] = payload.regional;
      linhaBebe[3] = normalizarFilialId(payload.filial);
      linhaBebe[4] = payload.nomeBebe;
      linhaBebe[5] = "Não Informado";
      linhaBebe[9] = formatarDataInclusao(payload.nascBebe);
      linhaBebe[10] = "BEBÊ LACTANTE";
      linhaBebe[11] = ""; // ID VAZIO (Ignora no Dashboard)
      linhaBebe[12] = payload.ddd;
      linhaBebe[13] = payload.celular;
      linhaBebe[18] = payload.emailMagalu;
      linhaBebe[20] = payload.tipoTransporte;
      linhaBebe[22] = payload.aeroporto1;
      linhaBebe[23] = payload.aeroporto2;
      linhaBebe[24] = payload.cidadeEmbarque;
      
      sheet.appendRow(linhaBebe);
    }
    
    return { sucesso: true, msg: "Dados validados e salvos com sucesso!" };
    
  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor ao salvar os dados: " + e.message };
  }
}
