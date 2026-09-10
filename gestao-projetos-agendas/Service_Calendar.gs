// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Service_Calendar.gs
// Responsabilidade: Integração resiliente com Google Calendar API
// =============================================================================

/**
 * Busca eventos do calendário principal do usuário (CalendarApp.getDefaultCalendar())
 * para o intervalo da semana em visualização (Segunda a Domingo).
 * 
 * Implementação resiliente com fallback seguro para não quebrar a aplicação caso
 * o usuário não tenha autorizado a permissão de calendário ou esteja em modo restrito.
 * 
 * @param {string} dataInicioISO Data inicial YYYY-MM-DD (Segunda-feira)
 * @param {string} dataFimISO Data final YYYY-MM-DD (Domingo)
 * @return {Array<Object>} Lista de eventos formatados para a UI
 */
function Service_Calendar_buscarEventosSemana(dataInicioISO, dataFimISO) {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    if (!cal) {
      console.warn('Calendário padrão do usuário não acessível.');
      return [];
    }
    
    // Converte datas para instâncias de Date considerando início e fim do dia
    const dtInicio = Utils_parseDataSegura(dataInicioISO) || new Date();
    dtInicio.setHours(0, 0, 0, 0);
    
    const dtFim = Utils_parseDataSegura(dataFimISO) || new Date();
    dtFim.setHours(23, 59, 59, 999);
    
    // Executa a busca nativa de eventos no Google Calendar
    const eventosBrutos = cal.getEvents(dtInicio, dtFim);
    const eventosFormatados = [];
    
    for (let i = 0; i < eventosBrutos.length; i++) {
      const ev = eventosBrutos[i];
      const isDiaInteiro = ev.isAllDayEvent();
      const inicio = ev.getStartTime();
      const fim = ev.getEndTime();
      
      const dataEventoISO = Utils_formatarDataISO(inicio);
      
      eventosFormatados.push({
        id: ev.getId(),
        titulo: ev.getTitle() || '(Sem Título)',
        dataISO: dataEventoISO,
        dataBR: Utils_formatarDataBR(dataEventoISO),
        horaInicio: isDiaInteiro ? 'Dia Inteiro' : Utils_formatarHora(inicio),
        horaFim: isDiaInteiro ? '' : Utils_formatarHora(fim),
        diaInteiro: isDiaInteiro,
        local: ev.getLocation() || '',
        descricao: ev.getDescription() ? ev.getDescription().substring(0, 200) : '',
        cor: ev.getColor() || '#7C3AED'
      });
    }
    
    // Ordena eventos cronologicamente
    eventosFormatados.sort((a, b) => {
      if (a.dataISO !== b.dataISO) {
        return a.dataISO.localeCompare(b.dataISO);
      }
      if (a.diaInteiro && !b.diaInteiro) return -1;
      if (!a.diaInteiro && b.diaInteiro) return 1;
      return a.horaInicio.localeCompare(b.horaInicio);
    });
    
    return eventosFormatados;
  } catch (error) {
    console.error('Aviso: Falha controlada ao ler eventos do Google Calendar:', error.message);
    // Retorna array vazio para preservar a navegação das tarefas da semana
    return [];
  }
}

/**
 * Cria opcionalmente um compromisso rápido na agenda principal
 * 
 * @param {Object} eventoPayload Dados do evento a ser criado
 * @return {Object} Confirmação
 */
function Service_Calendar_criarEventoRapido(eventoPayload) {
  try {
    const cal = CalendarApp.getDefaultCalendar();
    if (!cal) throw new Error('Calendário principal indisponível.');
    
    const titulo = Security_sanitizarString(eventoPayload.titulo || 'Compromisso GP360');
    const dataISO = eventoPayload.dataISO || Utils_obterDataHojeISO();
    const diaInteiro = eventoPayload.diaInteiro === true;
    
    let eventoCriado;
    const dt = Utils_parseDataSegura(dataISO);
    
    if (diaInteiro) {
      eventoCriado = cal.createAllDayEvent(titulo, dt, {
        description: eventoPayload.descricao || ''
      });
    } else {
      const horaInicioStr = eventoPayload.horaInicio || '09:00';
      const horaFimStr = eventoPayload.horaFim || '10:00';
      
      const partesIni = horaInicioStr.split(':');
      const partesFim = horaFimStr.split(':');
      
      const dtIni = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), parseInt(partesIni[0], 10), parseInt(partesIni[1], 10), 0);
      const dtFim = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), parseInt(partesFim[0], 10), parseInt(partesFim[1], 10), 0);
      
      eventoCriado = cal.createEvent(titulo, dtIni, dtFim, {
        description: eventoPayload.descricao || '',
        location: eventoPayload.local || ''
      });
    }
    
    return Utils_criarRespostaSucesso({
      id: eventoCriado.getId(),
      titulo: titulo,
      dataISO: dataISO
    }, 'Evento adicionado ao Google Calendar.');
  } catch (error) {
    return Utils_criarRespostaErro(error);
  }
}
