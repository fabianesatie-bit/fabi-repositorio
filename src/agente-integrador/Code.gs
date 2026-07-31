/**
 * Inicializa a aplicacao Web App renderizando o arquivo Index.html.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard de Performance & Feedback')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * VERSÃO FINAL AUDITADA: Implementação de Árvore Hierárquica (Global > Diretoria > Regional)
 * e Média Ponderada Financeira.
 * @return {string} JSON estruturado com a análise de impacto multinível.
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feedbackSheet = ss.getSheetByName('REGISTRO_FEEDBACK');
    
    if (!feedbackSheet) {
      throw new Error('Aba "REGISTRO_FEEDBACK" nao foi encontrada. Verifique o nome da aba.');
    }
    
    // Leitura em lote otimizada
    const allData = feedbackSheet.getDataRange().getValues();
    if (allData.length <= 1) {
       return JSON.stringify({}); 
    }

    const headers = allData.shift();
    
    const idx = {
      diretoria: headers.indexOf('DIRETORIA'),
      regional: headers.indexOf('REGIONAL'),
      familiaCargo: headers.indexOf('FAMILIA_CARGO'),
      diasContratados: headers.indexOf('DIAS_CONTRATADOS'),
      f15: headers.indexOf('FEEDBACK_15_DIAS'),
      f30: headers.indexOf('FEEDBACK_30_DIAS'),
      f45: headers.indexOf('FEEDBACK_45_DIAS'),
      venda1P: headers.indexOf('VENDA 1P'),
      meta1P: headers.indexOf('META 1P'),
      nome: headers.indexOf('NOME'),
      filial: headers.indexOf('FILIAL')
    };

    // Validação de colunas obrigatórias
    if (idx.regional === -1 || idx.venda1P === -1 || idx.meta1P === -1) {
      throw new Error('Colunas estruturais ausentes (REGIONAL, VENDA 1P ou META 1P). Verifique o cabeçalho.');
    }

    const hasDiretoria = idx.diretoria !== -1;

    // Estrutura de Armazenamento Hierárquico
    const dadosBrutos = {
      global: [],
      diretorias: {},
      regionais: {},
      mapa: {} // Indexador: Qual Regional pertence a qual Diretoria
    };

    // 1. Varredura Única e Distribuição nos Níveis
    allData.forEach(row => {
      const dias = Number(row[idx.diasContratados]) || 0;
      const cargo = String(row[idx.familiaCargo] || '').trim();
      
      if (cargo === 'Vendedor' && dias >= 15) {
        const regional = String(row[idx.regional] || 'Sem Regional').trim();
        const diretoria = hasDiretoria ? String(row[idx.diretoria] || 'Sem Diretoria').trim() : 'Geral';
        
        // Mapeamento de Cascata
        if (!dadosBrutos.mapa[diretoria]) {
          dadosBrutos.mapa[diretoria] = new Set();
        }
        dadosBrutos.mapa[diretoria].add(regional);

        // Distribuição de Dados
        dadosBrutos.global.push(row);
        
        if (!dadosBrutos.diretorias[diretoria]) dadosBrutos.diretorias[diretoria] = [];
        dadosBrutos.diretorias[diretoria].push(row);
        
        if (!dadosBrutos.regionais[regional]) dadosBrutos.regionais[regional] = [];
        dadosBrutos.regionais[regional].push(row);
      }
    });

    // 2. Motor de Processamento Financeiro (Reutilizável para os 3 níveis)
    const processarMetricas = (rows) => {
      let somaVendasEmDia = 0, somaMetasEmDia = 0;
      let somaVendasPendentes = 0, somaMetasPendentes = 0;
      let colaboradoresPendentes = [];

      rows.forEach(row => {
        const dias = Number(row[idx.diasContratados]);
        const f15 = Number(row[idx.f15]) === 1;
        const f30 = Number(row[idx.f30]) === 1;
        const f45 = Number(row[idx.f45]) === 1;
        const isPendente = (dias >= 15 && !f15) || (dias >= 30 && !f30) || (dias >= 45 && !f45);
        const venda = Number(row[idx.venda1P]) || 0;
        const meta = Number(row[idx.meta1P]) || 0;
        
        if (isPendente) {
          somaVendasPendentes += venda;
          somaMetasPendentes += meta;
          
          let pendenciasStr = [];
          let diasAtraso = 0;
          if (dias >= 15 && !f15) { pendenciasStr.push('FB 15 Dias'); diasAtraso = Math.max(diasAtraso, dias - 15); }
          if (dias >= 30 && !f30) { pendenciasStr.push('FB 30 Dias'); diasAtraso = Math.max(diasAtraso, dias - 30); }
          if (dias >= 45 && !f45) { pendenciasStr.push('FB 45 Dias'); diasAtraso = Math.max(diasAtraso, dias - 45); }
          
          colaboradoresPendentes.push({
            nome: String(row[idx.nome] || 'Sem Nome').trim(),
            filial: String(row[idx.filial] || 'Sem Filial').trim(),
            diasContratados: dias,
            quaisPendentes: pendenciasStr.join(', '),
            diasAtraso: diasAtraso,
            atingimento: meta > 0 ? (venda / meta) * 100 : 0
          });
        } else {
          somaVendasEmDia += venda;
          somaMetasEmDia += meta;
        }
      });

      const mediaAtingimentoEmDia = somaMetasEmDia > 0 ? (somaVendasEmDia / somaMetasEmDia) * 100 : 0;
      const mediaAtingimentoPendente = somaMetasPendentes > 0 ? (somaVendasPendentes / somaMetasPendentes) * 100 : 0;
      
      const faturamentoPotencialPendentes = somaMetasPendentes * (mediaAtingimentoEmDia / 100);
      let faturamentoPerdido = faturamentoPotencialPendentes - somaVendasPendentes;
      if (faturamentoPerdido < 0) faturamentoPerdido = 0; 

      const somaVendasAtual = somaVendasEmDia + somaVendasPendentes;
      let incrementoPotencialPercentual = 0;
      if (somaVendasAtual > 0 && faturamentoPerdido > 0) {
          incrementoPotencialPercentual = (faturamentoPerdido / somaVendasAtual) * 100;
      }

      return {
        mediaAtingimentoEmDia: mediaAtingimentoEmDia,
        mediaAtingimentoPendente: mediaAtingimentoPendente,
        faturamentoPerdido: faturamentoPerdido,
        incrementoPotencialPercentual: incrementoPotencialPercentual,
        totalVendedores: rows.length,
        totalPendentes: colaboradoresPendentes.length,
        colaboradoresPendentes: colaboradoresPendentes.sort((a,b) => b.diasAtraso - a.diasAtraso)
      };
    };

    // Transformação Set para Array (Para serialização JSON segura)
    const mapToObjectArray = (map) => {
      const obj = {};
      for (const key in map) obj[key] = Array.from(map[key]);
      return obj;
    };

    // 3. Montagem do Payload Final
    const finalResult = {
      global: processarMetricas(dadosBrutos.global),
      diretorias: {},
      regionais: {},
      mapa: mapToObjectArray(dadosBrutos.mapa)
    };

    for (const dir in dadosBrutos.diretorias) {
      finalResult.diretorias[dir] = processarMetricas(dadosBrutos.diretorias[dir]);
    }
    
    for (const reg in dadosBrutos.regionais) {
      finalResult.regionais[reg] = processarMetricas(dadosBrutos.regionais[reg]);
    }

    return JSON.stringify(finalResult);

  } catch (error) {
    return JSON.stringify({
      _error: true,
      message: error.message
    });
  }
}
