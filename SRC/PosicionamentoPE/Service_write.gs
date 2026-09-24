// =====================================================================
// SERVIÇO DE GRAVAÇÃO E PERSISTÊNCIA DE DADOS - ECOSSISTEMA GP360
// MÓDULO: FORMULÁRIO DE INSCRIÇÃO (PROJETO: Base Posicionamento estrategico)
// ALINHAMENTO ESTRITO COM AS 35 COLUNAS (A até AI) DE 'Detalhamento Regional'
// =====================================================================

/**
 * GATILHO PRINCIPAL DE EXIBIÇÃO WEB (DO-GET SEM TRAVA TEMPORAL)
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inscrição - Posicionamento Estratégico 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------------
// REGRA GLOBAL OBRIGATÓRIA GP360: NORMALIZAÇÃO DE FILIAIS (MIDDLEWARE LOCAL)
// ---------------------------------------------------------------------
function normalizarFilialIdWrite(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * ALIAS DE SEGURANÇA PARA COMPATIBILIDADE COM O FRONTEND
 */
function salvarDados(payload) {
  return gravarInscricaoEmLote(payload);
}

/**
 * ROTINA DE GRAVAÇÃO DA INSCRIÇÃO NA PLANILHA DESTINO
 * Mapeamento de 35 colunas exatas (A até AI) na aba Detalhamento Regional
 * @param {Object} payload - Objeto contendo os dados do participante e acompanhantes
 */
function gravarInscricaoEmLote(payload) {
  try {
    if (!payload) {
      return { sucesso: false, msg: "Dados de inscrição incompletos ou inválidos." };
    }

    var idCracha = String(payload.id || (payload.colaborador ? payload.colaborador.id : "")).trim();
    var nomeCompleto = payload.nomeCompleto || (payload.colaborador ? payload.colaborador.nome : "");
    var regional = payload.regional || (payload.colaborador ? payload.colaborador.regional : "");
    var filial = payload.filial || (payload.colaborador ? payload.colaborador.filial : "");
    var rg = payload.rg || (payload.colaborador ? payload.colaborador.rg : "");
    var orgao = payload.orgao || (payload.colaborador ? payload.colaborador.orgao : "");
    var cpf = payload.cpf || (payload.colaborador ? payload.colaborador.cpf : "");
    var nascimento = payload.nascimento || (payload.colaborador ? payload.colaborador.nascimento : "");
    var cargo = payload.cargo || (payload.colaborador ? payload.colaborador.cargo : "");

    // 1. CONEXÃO COM A PLANILHA DESTINO
    var ssDestino = SpreadsheetApp.openById(SPREADSHEET_DESTINO_ID);
    var sheetDestino = ssDestino.getSheetByName('Detalhamento Regional');
    if (!sheetDestino) {
      return { sucesso: false, msg: "Aba 'Detalhamento Regional' não encontrada na planilha de destino." };
    }

    // 2. DOUBLE-LOCK NO SERVIDOR (ANTI-DUPLICIDADE NA COLUNA L - INDEX 11)
    var dadosDestino = sheetDestino.getDataRange().getValues();
    for (var k = 1; k < dadosDestino.length; k++) {
      var matriculaGravada = String(dadosDestino[k][11]).trim(); // Coluna L (Índice 11) - ID (CRACHÁ)
      if (idCracha !== "" && matriculaGravada === idCracha) {
        return { 
          sucesso: false, 
          msg: "Você já preencheu este formulário anteriormente. Para alterações de rotas e outros detalhes, envie um e-mail para a equipe de organização." 
        };
      }
    }

    // 3. MONTAGEM DO ARRAY DE 35 COLUNAS EXATAS (A até AI)
    var timestampAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    var filialNormalizada = normalizarFilialIdWrite(filial);
    var emailUsuario = Session.getActiveUser().getEmail() || "";

    var transporteFinal = payload.tipoTransporte || "Próprio";
    if (transporteFinal.indexOf('Próprio') !== -1) {
      transporteFinal = 'Próprio';
    }

    var isLactanteTexto = payload.isLactante ? "Sim" : "Não";

    var linhaExata = [
      timestampAtual,                             // Coluna A (Index 0)  - Timestamp
      idCracha,                                   // Coluna B (Index 1)  - QT / ID
      regional,                                   // Coluna C (Index 2)  - REGIONAL
      filialNormalizada,                          // Coluna D (Index 3)  - FILIAL
      nomeCompleto,                               // Coluna E (Index 4)  - NOME COMPLETO
      payload.sexo || "Não Informado",            // Coluna F (Index 5)  - SEXO
      rg,                                         // Coluna G (Index 6)  - RG
      orgao,                                      // Coluna H (Index 7)  - ÓRGÃO EXPEDIDOR
      cpf,                                        // Coluna I (Index 8)  - CPF
      nascimento,                                 // Coluna J (Index 9)  - DATA NASCIMENTO
      cargo,                                      // Coluna K (Index 10) - CARGO
      idCracha,                                   // Coluna L (Index 11) - ID (CRACHÁ) [Double-Lock Key]
      payload.ddd || "",                          // Coluna M (Index 12) - DDD CELULAR
      payload.celular || "",                      // Coluna N (Index 13) - Nº CELULAR
      "Não",                                      // Coluna O (Index 14) - Hospedagem no Evento?
      "",                                         // Coluna P (Index 15) - Caso positivo, Qual data?
      "",                                         // Coluna Q (Index 16) - Dupla de quarto
      payload.restricaoAlimentar || "Nenhuma",    // Coluna R (Index 17) - RESTRIÇÃO ALIMENTAR
      payload.emailMagalu || emailUsuario,        // Coluna S (Index 18) - E-MAIL MAGALU
      payload.emailPessoal || "",                 // Coluna T (Index 19) - E-MAIL PESSOAL
      transporteFinal,                            // Coluna U (Index 20) - TIPO TRANSPORTE
      payload.observacaoRotas || "",              // Coluna V (Index 21) - OBSERVAÇÃO ROTAS
      payload.aeroporto1 || "",                   // Coluna W (Index 22) - AEROPORTO PREFERÊNCIA 1
      payload.aeroporto2 || "",                   // Coluna X (Index 23) - AEROPORTO PREFERÊNCIA 2
      payload.cidadeEmbarque || "",               // Coluna Y (Index 24) - CIDADE EMBARQUE
      payload.mobilidadeReduzida || "Não",        // Coluna Z (Index 25) - MOBILIDADE REDUZIDA
      payload.observacoesGerais || "",            // Coluna AA (Index 26)- OBSERVAÇÕES GERAIS
      payload.contatoEmergencia || "",            // Coluna AB (Index 27)- Contato de Suporte / Emergência
      isLactanteTexto,                            // Coluna AC (Index 28)- MÃE LACTANTE? (Sim / Não)
      payload.nomeAcomp || "",                    // Coluna AD (Index 29)- NOME DO ACOMPANHANTE
      payload.cpfAcomp || "",                     // Coluna AE (Index 30)- CPF DO ACOMPANHANTE
      payload.rgAcomp || "",                      // Coluna AF (Index 31)- RG DO ACOMPANHANTE
      payload.nascAcomp || "",                    // Coluna AG (Index 32)- DATA NASC. ACOMPANHANTE
      payload.nomeBebe || "",                     // Coluna AH (Index 33)- NOME DO BEBÊ
      payload.nascBebe || ""                      // Coluna AI (Index 34)- DATA NASC. BEBÊ
    ];

    // 4. ESCRITA ATÔMICA
    sheetDestino.appendRow(linhaExata);

    return { sucesso: true, msg: "Inscrição realizada e gravada com sucesso!" };

  } catch (e) {
    return { sucesso: false, msg: "Falha técnica ao gravar inscrição: " + e.message };
  }
}
