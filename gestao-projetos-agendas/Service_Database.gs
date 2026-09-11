// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Service_Database.gs
// Responsabilidade: Operações de banco de dados em lote no Google Sheets e CRUD
// =============================================================================

/**
 * Obtém a instância da planilha configurada para o banco de dados.
 * Se estiver vinculado a uma planilha (Container-bound), utiliza-a diretamente.
 * Caso contrário, consulta o ID persistido nas propriedades do script ou cria uma nova.
 * 
 * @return {Spreadsheet} Instância ativa do Google Spreadsheet
 */
function DB_obterPlanilha() {
  // 1. Tenta obter a planilha do container ativo
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) {
    return ativa;
  }
  
  // 2. Consulta o ID gravado em ScriptProperties
  const props = PropertiesService.getScriptProperties();
  const idSalvo = props.getProperty(SPREADSHEET_ID_PROPERTY);
  if (idSalvo) {
    try {
      return SpreadsheetApp.openById(idSalvo);
    } catch (e) {
      console.warn('Planilha anterior inacessível com ID:', idSalvo);
    }
  }
  
  // 3. Cria uma nova planilha automaticamente via setupDatabase
  const resultadoSetup = setupDatabase();
  return SpreadsheetApp.openById(resultadoSetup.spreadsheetId);
}

/**
 * Rotina mandatória de inicialização automática do banco de dados (setupDatabase).
 * Cria/verifica as 4 abas estruturais com cabeçalhos padronizados, congelamento e formatação visual.
 * 
 * @param {string} [idPlanilhaAlvo] ID opcional de planilha existente para configurar
 * @return {Object} Resposta com ID e URL da planilha configurada
 */
function setupDatabase(idPlanilhaAlvo) {
  return Security_executarComLock(function() {
    let ss;
    
    if (idPlanilhaAlvo) {
      ss = SpreadsheetApp.openById(idPlanilhaAlvo);
    } else {
      const ativa = SpreadsheetApp.getActiveSpreadsheet();
      if (ativa) {
        ss = ativa;
      } else {
        const props = PropertiesService.getScriptProperties();
        const idSalvo = props.getProperty(SPREADSHEET_ID_PROPERTY);
        if (idSalvo) {
          try {
            ss = SpreadsheetApp.openById(idSalvo);
          } catch (e) {
            ss = SpreadsheetApp.create(DEFAULT_SPREADSHEET_NAME);
          }
        } else {
          ss = SpreadsheetApp.create(DEFAULT_SPREADSHEET_NAME);
        }
      }
    }
    
    // Salva o ID da planilha para consultas futuras do script
    PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
    
    const abasNecessarias = [
      { nome: SHEET_NAMES.PROJETOS, cabecalhos: SHEET_HEADERS.PROJETOS },
      { nome: SHEET_NAMES.TAREFAS, cabecalhos: SHEET_HEADERS.TAREFAS },
      { nome: SHEET_NAMES.HISTORICO_TEMPO, cabecalhos: SHEET_HEADERS.HISTORICO_TEMPO },
      { nome: SHEET_NAMES.CONFIGURACOES, cabecalhos: SHEET_HEADERS.CONFIGURACOES }
    ];
    
    abasNecessarias.forEach(abaInfo => {
      let sheet = ss.getSheetByName(abaInfo.nome);
      if (!sheet) {
        sheet = ss.insertSheet(abaInfo.nome);
      }
      
      // Verifica se a primeira linha já possui os cabeçalhos
      const ultimaLinha = sheet.getLastRow();
      const ultimaColuna = sheet.getLastColumn();
      
      let precisaEscreverCabecalho = false;
      if (ultimaLinha === 0 || ultimaColuna === 0) {
        precisaEscreverCabecalho = true;
      } else {
        const cabecalhoExistente = sheet.getRange(1, 1, 1, Math.min(ultimaColuna, abaInfo.cabecalhos.length)).getValues()[0];
        if (!cabecalhoExistente || cabecalhoExistente[0] !== abaInfo.cabecalhos[0]) {
          precisaEscreverCabecalho = true;
        }
      }
      
      if (precisaEscreverCabecalho) {
        const range = sheet.getRange(1, 1, 1, abaInfo.cabecalhos.length);
        range.setValues([abaInfo.cabecalhos]);
        range.setFontWeight('bold');
        range.setBackground('#7C3AED');
        range.setFontColor('#FFFFFF');
        range.setHorizontalAlignment('center');
        sheet.setFrozenRows(1);
        
        // Auto-ajuste de colunas
        for (let col = 1; col <= abaInfo.cabecalhos.length; col++) {
          sheet.setColumnWidth(col, 160);
        }
      }
    });
    
    // Remove aba padrão 'Página1' ou 'Sheet1' se vazia e houver mais abas
    const folhas = ss.getSheets();
    if (folhas.length > 4) {
      folhas.forEach(f => {
        const nome = f.getName();
        if ((nome === 'Página1' || nome === 'Sheet1') && f.getLastRow() === 0) {
          try { ss.deleteSheet(f); } catch (e) {}
        }
      });
    }
    
    // Insere sementes iniciais se o banco estiver 100% vazio para demonstração imediata
    DB_semearDadosIniciaisSeNecessario(ss);
    
    return {
      sucesso: true,
      spreadsheetId: ss.getId(),
      url: ss.getUrl(),
      mensagem: 'Banco de dados GP360 configurado com sucesso.'
    };
  }, 15000);
}

/**
 * Insere dados iniciais de demonstração caso as tabelas estejam zeradas
 */
function DB_semearDadosIniciaisSeNecessario(ss) {
  const sheetProjetos = ss.getSheetByName(SHEET_NAMES.PROJETOS);
  const sheetTarefas = ss.getSheetByName(SHEET_NAMES.TAREFAS);
  
  if (sheetProjetos && sheetProjetos.getLastRow() <= 1) {
    const hojeISO = Utils_obterDataHojeISO();
    const idProjeto1 = Utils_gerarId('PRJ');
    const idProjeto2 = Utils_gerarId('PRJ');
    
    sheetProjetos.appendRow([
      idProjeto1,
      'Expansão e Eficiência Operacional',
      'Operações & Logística',
      '#EDE9FE', // Lavanda
      STATUS_PROJETO.ATIVO,
      hojeISO,
      'Projeto estratégico de revisão dos fluxos de atendimento.'
    ]);
    
    sheetProjetos.appendRow([
      idProjeto2,
      'Transformação Digital & Automação',
      'Tecnologia',
      '#D1FAE5', // Menta
      STATUS_PROJETO.ATIVO,
      hojeISO,
      'Implementação de integrações com Google Workspace e APIs.'
    ]);
    
    if (sheetTarefas && sheetTarefas.getLastRow() <= 1) {
      const dataAmanha = Utils_adicionarDias(hojeISO, 1);
      const dataOntem = Utils_adicionarDias(hojeISO, -2);
      
      // Tarefa 1: Hoje - Alta Prioridade
      sheetTarefas.appendRow([
        Utils_gerarId('TSK'),
        idProjeto1,
        'Alinhar cronograma com gerentes regionais',
        'Validar disponibilidade da equipe e definir metas da semana.',
        hojeISO,
        PRIORIDADE_TAREFA.ALTA,
        STATUS_TAREFA.EM_ANDAMENTO,
        '',
        '',
        1800,
        '',
        new Date().toISOString(),
        new Date().toISOString()
      ]);
      
      // Tarefa 2: Hoje - Status Aguardando (para teste dos badges coloridos)
      sheetTarefas.appendRow([
        Utils_gerarId('TSK'),
        idProjeto2,
        'Revisar especificação técnica da API de Calendário',
        'Aguardando retorno do comitê de segurança sobre escopos OAuth.',
        hojeISO,
        PRIORIDADE_TAREFA.MEDIA,
        STATUS_TAREFA.AGUARDANDO,
        dataOntem, // 2 dias atrás -> faixa 1-2 dias (azul suave)
        'Aprovação formal do escopo de leitura do Google Calendar',
        0,
        '',
        new Date().toISOString(),
        new Date().toISOString()
      ]);
      
      // Tarefa 3: Amanhã
      sheetTarefas.appendRow([
        Utils_gerarId('TSK'),
        idProjeto1,
        'Consolidar indicadores de tempo da semana',
        'Exportar dados do HISTORICO_TEMPO e estruturar apresentação.',
        dataAmanha,
        PRIORIDADE_TAREFA.MEDIA,
        STATUS_TAREFA.PENDENTE,
        '',
        '',
        0,
        '',
        new Date().toISOString(),
        new Date().toISOString()
      ]);
    }
  }
}

/**
 * Operação agregada de alta performance: carrega tudo o que a SPA precisa em UMA ÚNICA chamada RPC.
 * Lê projetos, tarefas e eventos do Google Calendar para o modo selecionado (Dia, Semana ou Mês).
 * 
 * @param {string} [dataReferenciaISO] Data YYYY-MM-DD (padrão: hoje)
 * @param {string} [modoVisualizacao] 'dia', 'semana' ou 'mes' (padrão: 'semana')
 * @return {Object} Pacote completo de dados para a interface
 */
function obterDadosIniciaisApp(dataReferenciaISO, modoVisualizacao) {
  try {
    const ss = DB_obterPlanilha();
    const modo = modoVisualizacao || 'semana';
    
    // Calcula metadados temporais para os 3 modos de visualização
    const infoDia = Utils_obterIntervaloDia(dataReferenciaISO);
    const infoSemana = Utils_obterIntervaloSemana(dataReferenciaISO);
    const infoMes = Utils_obterIntervaloMes(dataReferenciaISO);
    const hojeISO = Utils_obterDataHojeISO();
    
    // 1. Leitura em lote de PROJETOS
    const sheetProjetos = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    const dadosProjetosBrutos = sheetProjetos ? sheetProjetos.getDataRange().getValues() : [];
    const projetos = [];
    
    for (let i = 1; i < dadosProjetosBrutos.length; i++) {
      const linha = dadosProjetosBrutos[i];
      if (!linha[0]) continue;
      
      projetos.push({
        ID_Projeto: String(linha[0]).trim(),
        Nome_Projeto: String(linha[1] || '').trim(),
        Area_Contexto: String(linha[2] || '').trim(),
        Cor_Pastel: String(linha[3] || '#EDE9FE').trim(),
        Status: String(linha[4] || STATUS_PROJETO.ATIVO).trim(),
        Data_Criacao: Utils_formatarDataISO(linha[5]),
        Observacoes: String(linha[6] || '').trim()
      });
    }
    
    // Mapeamento rápido de projetos por ID para associação com tarefas
    const mapaProjetos = {};
    projetos.forEach(p => { mapaProjetos[p.ID_Projeto] = p; });
    
    // 2. Leitura em lote de TAREFAS
    const sheetTarefas = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    const dadosTarefasBrutos = sheetTarefas ? sheetTarefas.getDataRange().getValues() : [];
    const tarefas = [];
    
    for (let i = 1; i < dadosTarefasBrutos.length; i++) {
      const linha = dadosTarefasBrutos[i];
      if (!linha[0]) continue;
      
      const idTarefa = String(linha[0]).trim();
      const idProjeto = String(linha[1] || '').trim();
      const status = String(linha[6] || STATUS_TAREFA.PENDENTE).trim();
      const dataProgramadaISO = Utils_formatarDataISO(linha[4]);
      const dataInicioAguardandoISO = Utils_formatarDataISO(linha[7]);
      const tempoTotalSegundos = parseInt(linha[9], 10) || 0;
      
      // Cálculo dinâmico do contador e do badge visual de "Dias Aguardando"
      let diasAguardando = 0;
      let badgeAguardando = null;
      
      if (status === STATUS_TAREFA.AGUARDANDO) {
        if (dataInicioAguardandoISO) {
          diasAguardando = Utils_calcularDiasDecorridos(dataInicioAguardandoISO, hojeISO);
        } else {
          diasAguardando = 1;
        }
        
        if (diasAguardando <= 2) {
          badgeAguardando = {
            dias: diasAguardando,
            faixa: '1-2',
            bg: FAIXAS_AGUARDANDO.FAIXA_1.bg,
            texto: FAIXAS_AGUARDANDO.FAIXA_1.texto,
            border: FAIXAS_AGUARDANDO.FAIXA_1.border,
            label: `${diasAguardando}d aguardando`
          };
        } else if (diasAguardando <= 5) {
          badgeAguardando = {
            dias: diasAguardando,
            faixa: '3-5',
            bg: FAIXAS_AGUARDANDO.FAIXA_2.bg,
            texto: FAIXAS_AGUARDANDO.FAIXA_2.texto,
            border: FAIXAS_AGUARDANDO.FAIXA_2.border,
            label: `${diasAguardando}d aguardando`
          };
        } else {
          badgeAguardando = {
            dias: diasAguardando,
            faixa: '6+',
            bg: FAIXAS_AGUARDANDO.FAIXA_3.bg,
            texto: FAIXAS_AGUARDANDO.FAIXA_3.texto,
            border: FAIXAS_AGUARDANDO.FAIXA_3.border,
            label: `${diasAguardando}d aguardando!`
          };
        }
      }
      
      const projetoAssociado = mapaProjetos[idProjeto] || null;
      
      tarefas.push({
        ID_Tarefa: idTarefa,
        ID_Projeto: idProjeto,
        Nome_Projeto: projetoAssociado ? projetoAssociado.Nome_Projeto : 'Sem Projeto',
        Cor_Projeto: projetoAssociado ? projetoAssociado.Cor_Pastel : '#EDE9FE',
        Titulo: String(linha[2] || '').trim(),
        Proximos_Passos: String(linha[3] || '').trim(),
        Data_Programada: dataProgramadaISO,
        Data_Programada_BR: Utils_formatarDataBR(dataProgramadaISO),
        Prioridade: String(linha[5] || PRIORIDADE_TAREFA.MEDIA).trim(),
        Status: status,
        Data_Inicio_Aguardando: dataInicioAguardandoISO,
        Motivo_Aguardando: String(linha[8] || '').trim(),
        Tempo_Total_Segundos: tempoTotalSegundos,
        Tempo_Total_Formatado: Utils_formatarSegundosParaHHMMSS(tempoTotalSegundos),
        Data_Conclusao: Utils_formatarDataISO(linha[10]),
        Criado_Em: String(linha[11] || ''),
        Atualizado_Em: String(linha[12] || ''),
        diasAguardando: diasAguardando,
        badgeAguardando: badgeAguardando
      });
    }
    
    // 3. Leitura dos eventos do Google Calendar para o intervalo adequado ao modo
    let eventosCalendario = [];
    try {
      let dataInicioBusca = infoSemana.segundaISO;
      let dataFimBusca = infoSemana.domingoISO;
      
      if (modo === 'dia') {
        dataInicioBusca = infoDia.dataISO;
        dataFimBusca = infoDia.dataISO;
      } else if (modo === 'mes') {
        dataInicioBusca = infoMes.inicioGridISO;
        dataFimBusca = infoMes.fimGridISO;
      }
      
      eventosCalendario = Service_Calendar_buscarEventosSemana(dataInicioBusca, dataFimBusca);
    } catch (eCal) {
      console.warn('Falha ao obter eventos do calendário:', eCal);
    }
    
    return Utils_criarRespostaSucesso({
      modoVisualizacao: modo,
      infoDia: infoDia,
      infoSemana: infoSemana,
      infoMes: infoMes,
      projetos: projetos,
      tarefas: tarefas,
      eventosCalendario: eventosCalendario,
      usuarioLogado: Security_obterEmailUsuarioAtivo(),
      hojeISO: hojeISO,
      paletaCores: PALETA_CORES_PASTEL
    }, 'Dados sincronizados com sucesso.');
  } catch (error) {
    console.error('Erro em obterDadosIniciaisApp:', error);
    return Utils_criarRespostaErro(error);
  }
}

/**
 * Cria ou atualiza uma tarefa com bloqueio de concorrência e controle de estados
 * 
 * @param {Object} tarefaPayload Dados da tarefa enviados pelo frontend
 * @return {Object} Resposta com tarefa salva
 */
function salvarTarefa(tarefaPayload) {
  return Security_executarComLock(function() {
    const payload = Security_validarPayloadTarefa(tarefaPayload);
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    if (!sheet) throw new Error('Aba TAREFAS não localizada no banco de dados.');
    
    const hojeISO = Utils_obterDataHojeISO();
    const agoraISO = new Date().toISOString();
    
    const dados = sheet.getDataRange().getValues();
    let linhaEncontrada = -1;
    let dadosLinhaAntiga = null;
    
    if (payload.ID_Tarefa) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]).trim() === payload.ID_Tarefa) {
          linhaEncontrada = i + 1;
          dadosLinhaAntiga = dados[i];
          break;
        }
      }
    }
    
    // Determina transições de status para Aguardando e Concluída
    let dataInicioAguardando = payload.Data_Inicio_Aguardando;
    let motivoAguardando = payload.Motivo_Aguardando;
    let dataConclusao = '';
    let tempoTotal = payload.Tempo_Total_Segundos;
    let criadoEm = agoraISO;
    
    if (linhaEncontrada > 0 && dadosLinhaAntiga) {
      const statusAnterior = String(dadosLinhaAntiga[6] || '').trim();
      criadoEm = dadosLinhaAntiga[11] || agoraISO;
      tempoTotal = parseInt(dadosLinhaAntiga[9], 10) || tempoTotal;
      dataConclusao = dadosLinhaAntiga[10] ? Utils_formatarDataISO(dadosLinhaAntiga[10]) : '';
      
      // Se acabou de virar Aguardando
      if (payload.Status === STATUS_TAREFA.AGUARDANDO) {
        if (statusAnterior !== STATUS_TAREFA.AGUARDANDO || !dataInicioAguardando) {
          dataInicioAguardando = hojeISO;
        }
      } else {
        // Saiu de Aguardando: limpa o contador
        dataInicioAguardando = '';
        motivoAguardando = '';
      }
      
      // Se acabou de virar Concluída
      if (payload.Status === STATUS_TAREFA.CONCLUIDA) {
        if (statusAnterior !== STATUS_TAREFA.CONCLUIDA || !dataConclusao) {
          dataConclusao = hojeISO;
        }
      } else {
        dataConclusao = '';
      }
    } else {
      // Nova tarefa
      payload.ID_Tarefa = Utils_gerarId('TSK');
      if (payload.Status === STATUS_TAREFA.AGUARDANDO) {
        dataInicioAguardando = hojeISO;
      }
      if (payload.Status === STATUS_TAREFA.CONCLUIDA) {
        dataConclusao = hojeISO;
      }
    }
    
    const linhaGravar = [
      payload.ID_Tarefa,
      payload.ID_Projeto,
      payload.Titulo,
      payload.Proximos_Passos,
      payload.Data_Programada || hojeISO,
      payload.Prioridade,
      payload.Status,
      dataInicioAguardando,
      motivoAguardando,
      tempoTotal,
      dataConclusao,
      criadoEm,
      agoraISO
    ];
    
    if (linhaEncontrada > 0) {
      sheet.getRange(linhaEncontrada, 1, 1, linhaGravar.length).setValues([linhaGravar]);
    } else {
      sheet.appendRow(linhaGravar);
    }
    
    return Utils_criarRespostaSucesso({
      ID_Tarefa: payload.ID_Tarefa,
      Status: payload.Status,
      Data_Programada: payload.Data_Programada || hojeISO
    }, 'Tarefa gravada com sucesso.');
  });
}

/**
 * Atualização rápida de status via card da tarefa
 * 
 * @param {string} idTarefa ID da tarefa
 * @param {string} novoStatus Novo status desejado
 * @param {string} [motivoAguardando] Motivo caso o status seja 'Aguardando'
 * @return {Object} Confirmação da operação
 */
function atualizarStatusTarefa(idTarefa, novoStatus, motivoAguardando) {
  return Security_executarComLock(function() {
    if (!idTarefa) throw new Error('ID da tarefa não informado.');
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    if (!sheet) throw new Error('Aba TAREFAS não localizada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaIndex = -1;
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idTarefa).trim()) {
        linhaIndex = i + 1;
        break;
      }
    }
    
    if (linhaIndex === -1) {
      throw new Error(`Tarefa com ID ${idTarefa} não foi encontrada.`);
    }
    
    const hojeISO = Utils_obterDataHojeISO();
    const agoraISO = new Date().toISOString();
    
    let dataInicioAguardando = '';
    let motivoFinal = '';
    let dataConclusao = '';
    
    if (novoStatus === STATUS_TAREFA.AGUARDANDO) {
      dataInicioAguardando = hojeISO;
      motivoFinal = motivoAguardando ? Security_sanitizarString(motivoAguardando) : 'Aguardando ação externa';
    } else if (novoStatus === STATUS_TAREFA.CONCLUIDA) {
      dataConclusao = hojeISO;
    }
    
    // Coluna 7: Status | Coluna 8: Data_Inicio_Aguardando | Coluna 9: Motivo_Aguardando | Coluna 11: Data_Conclusao | Coluna 13: Atualizado_Em
    sheet.getRange(linhaIndex, 7).setValue(novoStatus);
    sheet.getRange(linhaIndex, 8).setValue(dataInicioAguardando);
    sheet.getRange(linhaIndex, 9).setValue(motivoFinal);
    sheet.getRange(linhaIndex, 11).setValue(dataConclusao);
    sheet.getRange(linhaIndex, 13).setValue(agoraISO);
    
    return Utils_criarRespostaSucesso({
      ID_Tarefa: idTarefa,
      Status: novoStatus,
      Data_Inicio_Aguardando: dataInicioAguardando,
      Motivo_Aguardando: motivoFinal
    }, `Status atualizado para "${novoStatus}".`);
  });
}

/**
 * Ação Rápida Manual de 1 Toque: Adia a tarefa programada para o dia seguinte (+1 dia)
 * 
 * @param {string} idTarefa ID da tarefa a ser adiada
 * @return {Object} Nova data programada e confirmação
 */
function adiarTarefaMaisUmDia(idTarefa) {
  return Security_executarComLock(function() {
    if (!idTarefa) throw new Error('ID da tarefa não informado.');
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    if (!sheet) throw new Error('Aba TAREFAS não localizada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaIndex = -1;
    let dataProgramadaAtual = '';
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idTarefa).trim()) {
        linhaIndex = i + 1;
        dataProgramadaAtual = Utils_formatarDataISO(dados[i][4]);
        break;
      }
    }
    
    if (linhaIndex === -1) {
      throw new Error(`Tarefa ${idTarefa} não localizada.`);
    }
    
    const baseCalculo = dataProgramadaAtual || Utils_obterDataHojeISO();
    const novaDataProgramada = Utils_adicionarDias(baseCalculo, 1);
    const agoraISO = new Date().toISOString();
    
    // Atualiza a coluna 5 (Data_Programada) e a coluna 13 (Atualizado_Em)
    sheet.getRange(linhaIndex, 5).setValue(novaDataProgramada);
    sheet.getRange(linhaIndex, 13).setValue(agoraISO);
    
    return Utils_criarRespostaSucesso({
      ID_Tarefa: idTarefa,
      Nova_Data_Programada: novaDataProgramada,
      Nova_Data_Programada_BR: Utils_formatarDataBR(novaDataProgramada)
    }, `Tarefa adiada para amanhã (${Utils_formatarDataBR(novaDataProgramada)}).`);
  });
}

/**
 * Exclui permanentemente uma tarefa do banco de dados
 * 
 * @param {string} idTarefa ID da tarefa
 * @return {Object} Confirmação
 */
function excluirTarefa(idTarefa) {
  return Security_executarComLock(function() {
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    if (!sheet) throw new Error('Aba TAREFAS não encontrada.');
    
    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idTarefa).trim()) {
        sheet.deleteRow(i + 1);
        return Utils_criarRespostaSucesso({ ID_Tarefa: idTarefa }, 'Tarefa removida com sucesso.');
      }
    }
    throw new Error('Tarefa não encontrada.');
  });
}

/**
 * Salva ou atualiza um projeto
 * 
 * @param {Object} projetoPayload Dados do projeto
 * @return {Object} Confirmação e ID do projeto
 */
function salvarProjeto(projetoPayload) {
  return Security_executarComLock(function() {
    const payload = Security_validarPayloadProjeto(projetoPayload);
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    if (!sheet) throw new Error('Aba PROJETOS não encontrada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaEncontrada = -1;
    let dataCriacao = Utils_obterDataHojeISO();
    
    if (payload.ID_Projeto) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]).trim() === payload.ID_Projeto) {
          linhaEncontrada = i + 1;
          dataCriacao = Utils_formatarDataISO(dados[i][5]) || dataCriacao;
          break;
        }
      }
    } else {
      payload.ID_Projeto = Utils_gerarId('PRJ');
    }
    
    const linhaGravar = [
      payload.ID_Projeto,
      payload.Nome_Projeto,
      payload.Area_Contexto,
      payload.Cor_Pastel,
      payload.Status,
      dataCriacao,
      payload.Observacoes
    ];
    
    if (linhaEncontrada > 0) {
      sheet.getRange(linhaEncontrada, 1, 1, linhaGravar.length).setValues([linhaGravar]);
    } else {
      sheet.appendRow(linhaGravar);
    }
    
    return Utils_criarRespostaSucesso(payload, 'Projeto salvo com sucesso.');
  });
}

/**
 * Registra uma sessão no HISTORICO_TEMPO e acumula a duração em TAREFAS
 * 
 * @param {string} idTarefa ID da tarefa vinculada
 * @param {string} dataExecucao Data no formato YYYY-MM-DD
 * @param {string} horaInicio Hora no formato HH:mm:ss
 * @param {string} horaFim Hora no formato HH:mm:ss
 * @param {number} duracaoSegundos Segundos transcorridos na sessão
 * @return {Object} Resposta com o novo total acumulado
 */
function registrarSessaoTempo(idTarefa, dataExecucao, horaInicio, horaFim, duracaoSegundos) {
  return Security_executarComLock(function() {
    if (!idTarefa) throw new Error('ID da tarefa obrigatório para registrar tempo.');
    const segs = Math.max(0, parseInt(duracaoSegundos, 10) || 0);
    if (segs <= 0) {
      return Utils_criarRespostaSucesso({}, 'Sessão com duração zerada descartada.');
    }
    
    const ss = DB_obterPlanilha();
    const sheetHistorico = ss.getSheetByName(SHEET_NAMES.HISTORICO_TEMPO);
    const sheetTarefas = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    
    if (!sheetHistorico || !sheetTarefas) {
      throw new Error('Abas de persistência de tempo não localizadas.');
    }
    
    const usuario = Security_obterEmailUsuarioAtivo();
    const dataReg = dataExecucao || Utils_obterDataHojeISO();
    const idSessao = Utils_gerarId('SES');
    
    // 1. Grava no HISTORICO_TEMPO
    sheetHistorico.appendRow([
      idSessao,
      idTarefa,
      dataReg,
      horaInicio || '',
      horaFim || '',
      segs,
      usuario
    ]);
    
    // 2. Acumula tempo total em TAREFAS
    const dadosTarefas = sheetTarefas.getDataRange().getValues();
    let novoTotal = segs;
    
    for (let i = 1; i < dadosTarefas.length; i++) {
      if (String(dadosTarefas[i][0]).trim() === String(idTarefa).trim()) {
        const tempoAnterior = parseInt(dadosTarefas[i][9], 10) || 0;
        novoTotal = tempoAnterior + segs;
        
        // Coluna 10: Tempo_Total_Segundos | Coluna 13: Atualizado_Em
        sheetTarefas.getRange(i + 1, 10).setValue(novoTotal);
        sheetTarefas.getRange(i + 1, 13).setValue(new Date().toISOString());
        break;
      }
    }
    
    return Utils_criarRespostaSucesso({
      idSessao: idSessao,
      idTarefa: idTarefa,
      duracaoSessaoSegundos: segs,
      tempoTotalSegundos: novoTotal,
      tempoTotalFormatado: Utils_formatarSegundosParaHHMMSS(novoTotal)
    }, `Sessão de ${Utils_formatarSegundosParaHHMMSS(segs)} registrada com sucesso.`);
  });
}

/**
 * Cria uma nova iniciativa / projeto com ID único e bloqueio de concorrência
 * 
 * @param {Object} projetoPayload Dados da iniciativa { Nome_Projeto, Area_Contexto, Cor_Pastel, Status, Observacoes }
 * @return {Object} Resposta com o projeto criado
 */
function criarProjeto(projetoPayload) {
  return Security_executarComLock(function() {
    if (!projetoPayload || !projetoPayload.Nome_Projeto) {
      throw new Error('O nome da iniciativa é obrigatório.');
    }
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    if (!sheet) throw new Error('Aba PROJETOS não localizada no banco de dados.');
    
    const idProjeto = Utils_gerarId('PRJ');
    const hojeISO = Utils_obterDataHojeISO();
    const nome = Security_sanitizarString(projetoPayload.Nome_Projeto);
    const area = Security_sanitizarString(projetoPayload.Area_Contexto || 'Geral');
    const cor = projetoPayload.Cor_Pastel || '#ffd7f6';
    const status = projetoPayload.Status || STATUS_PROJETO.ATIVO;
    const obs = Security_sanitizarString(projetoPayload.Observacoes || '');
    
    sheet.appendRow([
      idProjeto,
      nome,
      area,
      cor,
      status,
      hojeISO,
      obs
    ]);
    
    return Utils_criarRespostaSucesso({
      ID_Projeto: idProjeto,
      Nome_Projeto: nome,
      Area_Contexto: area,
      Cor_Pastel: cor,
      Status: status,
      Data_Criacao: hojeISO,
      Observacoes: obs
    }, `Iniciativa "${nome}" cadastrada com sucesso!`);
  });
}

/**
 * Atualiza os dados de uma iniciativa / projeto existente
 * 
 * @param {Object} projetoPayload Dados da iniciativa { ID_Projeto, Nome_Projeto, Area_Contexto, Cor_Pastel, Status, Observacoes }
 * @return {Object} Resposta com o projeto atualizado
 */
function atualizarProjeto(projetoPayload) {
  return Security_executarComLock(function() {
    if (!projetoPayload || !projetoPayload.ID_Projeto) {
      throw new Error('ID da iniciativa obrigatório para atualização.');
    }
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    if (!sheet) throw new Error('Aba PROJETOS não localizada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaIndex = -1;
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(projetoPayload.ID_Projeto).trim()) {
        linhaIndex = i + 1;
        break;
      }
    }
    
    if (linhaIndex === -1) {
      throw new Error(`Iniciativa ${projetoPayload.ID_Projeto} não localizada.`);
    }
    
    const nome = Security_sanitizarString(projetoPayload.Nome_Projeto);
    const area = Security_sanitizarString(projetoPayload.Area_Contexto || 'Geral');
    const cor = projetoPayload.Cor_Pastel || '#ffd7f6';
    const status = projetoPayload.Status || STATUS_PROJETO.ATIVO;
    const obs = Security_sanitizarString(projetoPayload.Observacoes || '');
    
    // Atualiza colunas 2 a 5 e 7 (não altera ID nem Data_Criacao)
    sheet.getRange(linhaIndex, 2).setValue(nome);
    sheet.getRange(linhaIndex, 3).setValue(area);
    sheet.getRange(linhaIndex, 4).setValue(cor);
    sheet.getRange(linhaIndex, 5).setValue(status);
    sheet.getRange(linhaIndex, 7).setValue(obs);
    
    return Utils_criarRespostaSucesso({
      ID_Projeto: projetoPayload.ID_Projeto,
      Nome_Projeto: nome,
      Area_Contexto: area,
      Cor_Pastel: cor,
      Status: status,
      Observacoes: obs
    }, `Iniciativa "${nome}" atualizada com sucesso!`);
  });
}

/**
 * Altera o status do projeto para Arquivado
 * 
 * @param {string} idProjeto ID do projeto a ser arquivado
 * @return {Object} Resposta com confirmação
 */
function arquivarProjeto(idProjeto) {
  return Security_executarComLock(function() {
    if (!idProjeto) throw new Error('ID do projeto não fornecido.');
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    if (!sheet) throw new Error('Aba PROJETOS não localizada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaIndex = -1;
    let nomeProjeto = '';
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idProjeto).trim()) {
        linhaIndex = i + 1;
        nomeProjeto = dados[i][1];
        break;
      }
    }
    
    if (linhaIndex === -1) {
      throw new Error(`Projeto ${idProjeto} não localizado.`);
    }
    
    // Coluna 5 é Status
    sheet.getRange(linhaIndex, 5).setValue(STATUS_PROJETO.ARQUIVADO);
    
    return Utils_criarRespostaSucesso({
      ID_Projeto: idProjeto,
      Status: STATUS_PROJETO.ARQUIVADO
    }, `Projeto "${nomeProjeto}" arquivado com sucesso.`);
  });
}

/**
 * Reativa/Desarquiva um projeto, retornando seu status para Ativo
 * 
 * @param {string} idProjeto ID do projeto a ser reativado
 * @return {Object} Resposta com confirmação
 */
function desarquivarProjeto(idProjeto) {
  return Security_executarComLock(function() {
    if (!idProjeto) throw new Error('ID do projeto não fornecido.');
    
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    if (!sheet) throw new Error('Aba PROJETOS não localizada.');
    
    const dados = sheet.getDataRange().getValues();
    let linhaIndex = -1;
    let nomeProjeto = '';
    
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idProjeto).trim()) {
        linhaIndex = i + 1;
        nomeProjeto = dados[i][1];
        break;
      }
    }
    
    if (linhaIndex === -1) {
      throw new Error(`Projeto ${idProjeto} não localizado.`);
    }
    
    // Coluna 5 é Status
    sheet.getRange(linhaIndex, 5).setValue(STATUS_PROJETO.ATIVO);
    
    return Utils_criarRespostaSucesso({
      ID_Projeto: idProjeto,
      Status: STATUS_PROJETO.ATIVO
    }, `Projeto "${nomeProjeto}" reativado para Ativo com sucesso.`);
  });
}

/**
 * Altera o status da tarefa para Arquivada
 * 
 * @param {string} idTarefa ID da tarefa a ser arquivada
 * @return {Object} Resposta com confirmação
 */
function arquivarTarefa(idTarefa) {
  return atualizarStatusTarefa(idTarefa, STATUS_TAREFA.ARQUIVADA);
}

/**
 * Restaura uma tarefa arquivada para o status Pendente
 * 
 * @param {string} idTarefa ID da tarefa a ser restaurada
 * @return {Object} Resposta com confirmação
 */
function desarquivarTarefa(idTarefa) {
  return atualizarStatusTarefa(idTarefa, STATUS_TAREFA.PENDENTE);
}

/**
 * Retorna o log completo de sessões gravadas em HISTORICO_TEMPO enriquecido
 * com dados do projeto e título da tarefa para a Visão de Governança e Métricas.
 * 
 * @return {Object} Lista de sessões
 */
function obterHistoricoTempoCompleto() {
  try {
    const ss = DB_obterPlanilha();
    const sheetHistorico = ss.getSheetByName(SHEET_NAMES.HISTORICO_TEMPO);
    const sheetTarefas = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    const sheetProjetos = ss.getSheetByName(SHEET_NAMES.PROJETOS);
    
    if (!sheetHistorico) return Utils_criarRespostaSucesso([]);
    
    const dadosHist = sheetHistorico.getDataRange().getValues();
    const dadosTarefas = sheetTarefas ? sheetTarefas.getDataRange().getValues() : [];
    const dadosProjetos = sheetProjetos ? sheetProjetos.getDataRange().getValues() : [];
    
    const mapaProjetos = {};
    for (let i = 1; i < dadosProjetos.length; i++) {
      if (dadosProjetos[i][0]) {
        mapaProjetos[String(dadosProjetos[i][0]).trim()] = {
          nome: String(dadosProjetos[i][1] || '').trim(),
          cor: String(dadosProjetos[i][3] || '#EDE9FE').trim()
        };
      }
    }
    
    const mapaTarefas = {};
    for (let i = 1; i < dadosTarefas.length; i++) {
      if (dadosTarefas[i][0]) {
        const idProj = String(dadosTarefas[i][1] || '').trim();
        const proj = mapaProjetos[idProj] || { nome: 'Geral', cor: '#EDE9FE' };
        mapaTarefas[String(dadosTarefas[i][0]).trim()] = {
          titulo: String(dadosTarefas[i][2] || '').trim(),
          idProjeto: idProj,
          nomeProjeto: proj.nome,
          corProjeto: proj.cor
        };
      }
    }
    
    const sessoes = [];
    for (let i = dadosHist.length - 1; i >= 1; i--) {
      const l = dadosHist[i];
      if (!l[0]) continue;
      
      const idTarefa = String(l[1] || '').trim();
      const tInfo = mapaTarefas[idTarefa] || { titulo: 'Tarefa Avulsa', nomeProjeto: 'Sem Projeto', corProjeto: '#EDE9FE' };
      const duracaoSegs = parseInt(l[5], 10) || 0;
      
      sessoes.push({
        idSessao: String(l[0]).trim(),
        idTarefa: idTarefa,
        tituloTarefa: tInfo.titulo,
        nomeProjeto: tInfo.nomeProjeto,
        corProjeto: tInfo.corProjeto,
        dataExecucao: Utils_formatarDataISO(l[2]),
        dataExecucaoBR: Utils_formatarDataBR(l[2]),
        horaInicio: String(l[3] || '').trim(),
        horaFim: String(l[4] || '').trim(),
        duracaoSegundos: duracaoSegs,
        duracaoFormatada: Utils_formatarSegundosParaHHMMSS(duracaoSegs),
        usuario: String(l[6] || '').trim()
      });
    }
    
    return Utils_criarRespostaSucesso(sessoes);
  } catch (err) {
    return Utils_criarRespostaErro('Falha ao obter histórico de tempo: ' + err.message);
  }
}

/**
 * Envia um e-mail formal de cobrança/follow-up de bloqueio via GmailApp
 * 
 * @param {string} destinatario E-mail de destino
 * @param {string} assunto Assunto do e-mail
 * @param {string} corpoTexto Mensagem em texto simples
 * @return {Object} Confirmação
 */
function enviarCobrancaBloqueio(destinatario, assunto, corpoTexto) {
  try {
    if (!destinatario || !assunto || !corpoTexto) {
      throw new Error('Destinatário, assunto e corpo são obrigatórios.');
    }
    GmailApp.sendEmail(destinatario, assunto, corpoTexto);
    return Utils_criarRespostaSucesso({}, 'E-mail de cobrança enviado com sucesso!');
  } catch (err) {
    return Utils_criarRespostaErro('Não foi possível enviar o e-mail: ' + err.message);
  }
}


