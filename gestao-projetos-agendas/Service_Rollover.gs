// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Service_Rollover.gs
// Responsabilidade: Rotina e instalador do gatilho automático de rollover diário (00:01)
// =============================================================================

/**
 * Função executada automaticamente pelo gatilho diário de madrugada (00:01).
 * Varre a tabela TAREFAS e identifica tarefas não concluídas com Data_Programada < Hoje.
 * Atualiza atomicamente a Data_Programada de todas elas para a data de Hoje.
 * 
 * @return {Object} Relatório com a quantidade de tarefas atualizadas
 */
function executarRolloverTarefas() {
  return Security_executarComLock(function() {
    const ss = DB_obterPlanilha();
    const sheet = ss.getSheetByName(SHEET_NAMES.TAREFAS);
    if (!sheet) {
      console.error('Aba TAREFAS não localizada para o Rollover.');
      return Utils_criarRespostaErro('Aba TAREFAS não localizada.');
    }
    
    const range = sheet.getDataRange();
    const valores = range.getValues();
    if (valores.length <= 1) {
      return Utils_criarRespostaSucesso({ totalAtualizadas: 0 }, 'Nenhuma tarefa cadastrada para rollover.');
    }
    
    const hojeISO = Utils_obterDataHojeISO();
    const agoraISO = new Date().toISOString();
    let tarefasAtualizadasCount = 0;
    const tarefasModificadas = [];
    
    // Varredura a partir da linha 2 (índice 1)
    for (let i = 1; i < valores.length; i++) {
      const linha = valores[i];
      const idTarefa = String(linha[0] || '').trim();
      if (!idTarefa) continue;
      
      const status = String(linha[6] || '').trim();
      const dataProgramada = Utils_formatarDataISO(linha[4]);
      
      // Condição de Rollover: Tarefa NÃO concluída e com data programada anterior a hoje
      if (status !== STATUS_TAREFA.CONCLUIDA && dataProgramada && dataProgramada < hojeISO) {
        // Atualiza a coluna 5 (Data_Programada) e coluna 13 (Atualizado_Em) na matriz em memória
        linha[4] = hojeISO;
        linha[12] = agoraISO;
        
        tarefasAtualizadasCount++;
        tarefasModificadas.push({
          id: idTarefa,
          titulo: linha[2],
          dataAnterior: dataProgramada,
          novaData: hojeISO
        });
      }
    }
    
    // Se houve alterações, grava a matriz completa de volta em UMA ÚNICA chamada em lote
    if (tarefasAtualizadasCount > 0) {
      range.setValues(valores);
      console.log(`[ROLLOVER SUCESSO] ${tarefasAtualizadasCount} tarefas atualizadas para a data de hoje (${hojeISO}).`);
    } else {
      console.log(`[ROLLOVER] Nenhuma tarefa pendente com data retroativa encontrada em ${hojeISO}.`);
    }
    
    return Utils_criarRespostaSucesso({
      hojeISO: hojeISO,
      tarefasAtualizadas: tarefasAtualizadasCount,
      detalhes: tarefasModificadas
    }, `${tarefasAtualizadasCount} tarefas tiveram sua data atualizada para hoje.`);
  }, 20000);
}

/**
 * Instala programaticamente o Gatilho Baseado em Tempo (Time-Driven Trigger)
 * para disparar a função executarRolloverTarefas diariamente na madrugada (00:00 - 01:00).
 * Remove previamente quaisquer gatilhos duplicados existentes da mesma função.
 * 
 * @return {Object} Status da instalação do gatilho
 */
function instalarGatilhoRollover() {
  try {
    const nomeFuncao = 'executarRolloverTarefas';
    const gatilhos = ScriptApp.getProjectTriggers();
    
    // Remove gatilhos anteriores da mesma função para evitar execuções duplicadas
    gatilhos.forEach(g => {
      if (g.getHandlerFunction() === nomeFuncao) {
        ScriptApp.deleteTrigger(g);
      }
    });
    
    // Cria novo gatilho diário programado para a madrugada (hora 0 = entre 00:00 e 01:00)
    ScriptApp.newTrigger(nomeFuncao)
      .timeBased()
      .atHour(0)
      .everyDays(1)
      .create();
      
    console.log('Gatilho de rollover diário configurado com sucesso.');
    return Utils_criarRespostaSucesso({
      funcao: nomeFuncao,
      frequencia: 'Diário (entre 00:00 e 01:00)'
    }, 'Gatilho de Rollover automático configurado com sucesso para a madrugada.');
  } catch (error) {
    console.error('Falha ao instalar gatilho de rollover:', error);
    return Utils_criarRespostaErro(error);
  }
}

/**
 * Remove o gatilho automático de rollover se solicitado pelo administrador
 * 
 * @return {Object} Confirmação de remoção
 */
function desinstalarGatilhoRollover() {
  try {
    const nomeFuncao = 'executarRolloverTarefas';
    const gatilhos = ScriptApp.getProjectTriggers();
    let removidos = 0;
    
    gatilhos.forEach(g => {
      if (g.getHandlerFunction() === nomeFuncao) {
        ScriptApp.deleteTrigger(g);
        removidos++;
      }
    });
    
    return Utils_criarRespostaSucesso({ removidos: removidos }, `${removidos} gatilho(s) de rollover removido(s).`);
  } catch (error) {
    return Utils_criarRespostaErro(error);
  }
}

/**
 * Consulta se o gatilho automático de rollover está ativo no projeto
 * 
 * @return {Object} Estado atual do gatilho
 */
function verificarStatusGatilhoRollover() {
  try {
    const nomeFuncao = 'executarRolloverTarefas';
    const gatilhos = ScriptApp.getProjectTriggers();
    const ativo = gatilhos.some(g => g.getHandlerFunction() === nomeFuncao);
    
    return Utils_criarRespostaSucesso({
      ativo: ativo,
      totalGatilhosProjeto: gatilhos.length
    });
  } catch (error) {
    return Utils_criarRespostaErro(error);
  }
}
