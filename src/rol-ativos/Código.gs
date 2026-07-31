/**
 * @file Code.gs
 * @description Arquitetura backend otimizada para integração Google Sheets e frontend HTML.
 * Processamento lógico seguro, isolado de race conditions na leitura e iterador sem limites fixos.
 */

const CONFIG = {
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1InLKT3qmWxAv7N-U1tyoSW0tNI-Qc4LTW0vyza1oSg0/edit#gid=0',
  SHEET_NAME: 'Base_Rol'
};

/**
 * Função obrigatória para inicializar a interface Web App.
 * @returns {HtmlOutput} Renderização do arquivo Index.html configurado.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Consulta de Funcionários')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Função auxiliar para limpeza agressiva de strings (remove espaços duplos e invisíveis).
 */
function limparString(str) {
  if (!str) return "";
  return str.toString().trim().replace(/\s+/g, ' ');
}

/**
 * Extrai lista de filiais únicas para popular o elemento <datalist> do frontend.
 * @returns {Array<string>} Array ordenado de filiais exclusivas.
 */
function getFiliaisDistintas() {
  try {
    const ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error("Aba não encontrada.");

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];

    const headers = data[0].map(h => limparString(h).toLowerCase());
    const idxFilial = headers.indexOf('filial');
    
    if (idxFilial === -1) return [];

    const filiaisSet = new Set();
    
    // Varredura completa na memória (sem limite artificial)
    for (let i = 1; i < data.length; i++) {
      const val = limparString(data[i][idxFilial]);
      if (val !== "") {
        filiaisSet.add(val);
      }
    }

    return Array.from(filiaisSet).sort();

  } catch (error) {
    console.error("Erro na extração de filiais: " + error.message);
    return [];
  }
}

/**
 * Realiza a consulta otimizada na planilha.
 * @param {Object} queryParams - Filtros (filial, id, nome).
 * @returns {Array<Object>|Object} Array de resultados ou objeto de erro mapeado.
 */
function buscarDados(queryParams) {
  try {
    const ss = SpreadsheetApp.openByUrl(CONFIG.SPREADSHEET_URL);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Aba '${CONFIG.SHEET_NAME}' não encontrada.`);
    }

    // Carregamento integral para a memória RAM do V8
    const data = sheet.getDataRange().getDisplayValues();
    
    if (data.length <= 1) return [];

    const headers = data[0].map(h => limparString(h).toLowerCase());
    const rows = data.slice(1);

    const colIndices = {
      competencia: headers.indexOf('competencia'),
      id: headers.indexOf('id'),
      nome: headers.indexOf('nome'),
      filial: headers.indexOf('filial'),
      tempo_empresa: headers.indexOf('tempo_empresa'),
      regional_filiais: headers.indexOf('regional_filiais'),
      situacao: headers.indexOf('situacao'),
      cargo: headers.indexOf('cargo'),
      login: headers.indexOf('login')
    };

    if (colIndices.id === -1 || colIndices.nome === -1 || colIndices.filial === -1) {
      throw new Error("Colunas obrigatórias ('id', 'nome', 'filial') ausentes na matriz de dados.");
    }

    const qFilial = queryParams.filial ? limparString(queryParams.filial).toLowerCase() : null;
    const qId = queryParams.id ? limparString(queryParams.id).toLowerCase() : null;
    const qNome = queryParams.nome ? limparString(queryParams.nome).toLowerCase() : null;

    const resultados = [];
    
    // Laço itera sobre todos os registros disponíveis na aba
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      let match = true;

      // Correspondência Estrita para Filial
      if (qFilial) {
        const valFilial = limparString(row[colIndices.filial]).toLowerCase();
        if (valFilial !== qFilial) match = false;
      }

      // Correspondência Estrita para ID
      if (match && qId) {
        const valId = limparString(row[colIndices.id]).toLowerCase();
        if (valId !== qId) match = false;
      }

      // Correspondência Parcial (Elástica) para Nome
      if (match && qNome) {
        const valNome = limparString(row[colIndices.nome]).toLowerCase();
        if (!valNome.includes(qNome)) match = false;
      }

      if (match) {
        resultados.push({
          competencia: colIndices.competencia !== -1 ? row[colIndices.competencia] : "",
          id: row[colIndices.id],
          nome: row[colIndices.nome],
          filial: row[colIndices.filial],
          tempo_empresa: colIndices.tempo_empresa !== -1 ? row[colIndices.tempo_empresa] : "",
          regional_filiais: colIndices.regional_filiais !== -1 ? row[colIndices.regional_filiais] : "",
          situacao: colIndices.situacao !== -1 ? row[colIndices.situacao] : "",
          cargo: colIndices.cargo !== -1 ? row[colIndices.cargo] : "",
          login: colIndices.login !== -1 ? row[colIndices.login] : ""
        });
      }
    }

    return resultados;

  } catch (error) {
    console.error("Exceção capturada em buscarDados:", error);
    return { error: true, message: error.message };
  }
}
