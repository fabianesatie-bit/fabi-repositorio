// =====================================================================
// SERVIÇO DE GRAVAÇÃO E VALIDAÇÃO ANTI-DUPLICIDADE (MULTILOCAÇÃO)
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
    // BLINDAGEM DE INTEGRIDADE: Zera aeroportos se transporte não for aéreo.
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
    
    // Verificação Anti-Duplicidade do Titular (Coluna L - índice 11)
    for (var i = 1; i < dados.length; i++) {
      var matriculaExistente = String(dados[i][11]).trim(); 
      if (matriculaExistente === String(payload.id).trim()) {
        return { 
          sucesso: false, 
          msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas ou dados, envie um e-mail diretamente para a Agência LTN." 
        };
      }
    }
    
    var linhasParaGravar = [];
    
    // --------------------------------------------------------------------
    // LINHA 1: GERENTE / COLABORADOR TITULAR
    // --------------------------------------------------------------------
    // Agora usando array de 28 posições para abranger até a Coluna AB (Índice 27)
    var linhaTitular = new Array(28).fill("");
    
    linhaTitular[0] = ""; // Coluna A - VAZIA (Agência)
    linhaTitular[1] = ""; // Coluna B - VAZIA (Agência)
    linhaTitular[2] = payload.regional; // Coluna C
    linhaTitular[3] = normalizarFilialId(payload.filial); // Coluna D
    linhaTitular[4] = payload.nomeCompleto; // Coluna E (Já validado com RG pelo Front-end)
    linhaTitular[5] = payload.sexo; // Coluna F
    linhaTitular[6] = payload.rg; // Coluna G
    linhaTitular[7] = payload.orgao; // Coluna H
    linhaTitular[8] = payload.cpf; // Coluna I
    
    var dataNasc = payload.nascimento;
    if (dataNasc && typeof dataNasc.getTime === 'function') {
      linhaTitular[9] = Utilities.formatDate(dataNasc, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      linhaTitular[9] = dataNasc; // Coluna J
    }
    
    linhaTitular[10] = payload.cargo; // Coluna K
    linhaTitular[11] = payload.id; // Coluna L - ID CRACHÁ (Chave Anti-duplicidade)
    linhaTitular[12] = payload.ddd; // Coluna M
    linhaTitular[13] = payload.celular; // Coluna N
    linhaTitular[14] = ""; // Coluna O
    linhaTitular[15] = ""; // Coluna P
    linhaTitular[16] = ""; // Coluna Q
    linhaTitular[17] = payload.restricaoAlimentar; // Coluna R
    linhaTitular[18] = payload.emailMagalu; // Coluna S
    linhaTitular[19] = payload.emailPessoal; // Coluna T
    linhaTitular[20] = payload.tipoTransporte; // Coluna U
    linhaTitular[21] = payload.observacaoRotas; // Coluna V
    linhaTitular[22] = payload.aeroporto1; // Coluna W
    linhaTitular[23] = payload.aeroporto2; // Coluna X
    linhaTitular[24] = payload.cidadeEmbarque; // Coluna Y
    linhaTitular[25] = payload.mobilidadeReduzida; // Coluna Z
    linhaTitular[26] = payload.observacoesGerais; // Coluna AA
    linhaTitular[27] = payload.contatoEmergencia; // Coluna AB - NOVO CAMPO DE EMERGÊNCIA
    
    linhasParaGravar.push(linhaTitular);
    
    // --------------------------------------------------------------------
    // LINHA 2 e 3: GRAVAÇÃO EM LOTE PARA LACTANTES E BEBÊ (Se aplicável)
    // --------------------------------------------------------------------
    if (payload.isLactante) {
      
      var linhaAcomp = new Array(28).fill("");
      linhaAcomp[2] = payload.regional; 
      linhaAcomp[3] = normalizarFilialId(payload.filial); 
      linhaAcomp[4] = payload.nomeAcomp; // Nome Acompanhante
      linhaAcomp[6] = payload.rgAcomp; // RG Acompanhante
      linhaAcomp[8] = payload.cpfAcomp; // CPF Acompanhante
      linhaAcomp[9] = payload.nascAcomp; // Nasc Acompanhante
      linhaAcomp[10] = "ACOMPANHANTE LACTANTE"; // Cargo especial
      linhaAcomp[11] = ""; // ID VAZIO (Exclui contagem no Dashboard)
      linhaAcomp[20] = payload.tipoTransporte; 
      linhaAcomp[22] = payload.aeroporto1; 
      linhaAcomp[23] = payload.aeroporto2; 
      linhaAcomp[24] = payload.cidadeEmbarque; 
      linhaAcomp[27] = payload.contatoEmergencia; // Herda o contato de emergência da titular
      
      var linhaBebe = new Array(28).fill("");
      linhaBebe[2] = payload.regional; 
      linhaBebe[3] = normalizarFilialId(payload.filial); 
      linhaBebe[4] = payload.nomeBebe; // Nome Bebê
      linhaBebe[9] = payload.nascBebe; // Nascimento Bebê
      linhaBebe[10] = "BEBÊ LACTANTE"; // Cargo especial
      linhaBebe[11] = ""; // ID VAZIO
      linhaBebe[20] = payload.tipoTransporte; 
      linhaBebe[22] = payload.aeroporto1; 
      linhaBebe[23] = payload.aeroporto2; 
      linhaBebe[24] = payload.cidadeEmbarque;
      
      linhasParaGravar.push(linhaAcomp);
      linhasParaGravar.push(linhaBebe);
    }
    
    // Efetua a escrita em lote (Multi-row append)
    for (var j = 0; j < linhasParaGravar.length; j++) {
       sheet.appendRow(linhasParaGravar[j]);
    }
    
    return { sucesso: true, msg: "Dados validados e salvos com sucesso!" };
    
  } catch (e) {
    return { sucesso: false, msg: "Erro no servidor ao salvar os dados: " + e.message };
  }
}
