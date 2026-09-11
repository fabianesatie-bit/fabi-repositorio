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
/**
 * Localiza especificamente o calendário 'Fabiane Satie' marcado na interface do Google Calendar.
 * 
 * Ordem estrita de identificação:
 * 1. Busca direta via CalendarApp.getCalendarsByName('Fabiane Satie') priorizando a que tem isSelected() === true.
 * 2. Busca em CalendarApp.getAllCalendars() pelo nome 'Fabiane Satie' ou contendo 'Fabiane'.
 * 3. Procura o calendário ativo com isSelected() === true que não seja de feriados/aniversários/tarefas/arena.
 * 4. Fallback final para CalendarApp.getDefaultCalendar().
 * 
 * @return {Calendar|null} O objeto Calendar correspondente à Fabiane Satie
 */
function Service_Calendar_obterCalendarioAlvo() {
  const nomeAlvo = (typeof CALENDARIO_CONFIG !== 'undefined' && CALENDARIO_CONFIG.NOME_AGENDA_ALVO)
    ? CALENDARIO_CONFIG.NOME_AGENDA_ALVO
    : 'Fabiane Satie';
  const nomeAlvoLower = nomeAlvo.trim().toLowerCase();

  // 1. Busca por nome exato / getCalendarsByName
  try {
    const calsPorNome = CalendarApp.getCalendarsByName(nomeAlvo);
    if (calsPorNome && calsPorNome.length > 0) {
      for (let i = 0; i < calsPorNome.length; i++) {
        const c = calsPorNome[i];
        try {
          if (typeof c.isSelected === 'function' && c.isSelected()) {
            console.log('Agenda Fabiane Satie localizada por nome e marcada na interface:', c.getName());
            return c;
          }
        } catch (eSel) {}
      }
      console.log('Agenda localizada por getCalendarsByName:', calsPorNome[0].getName());
      return calsPorNome[0];
    }
  } catch (eNome) {
    console.warn('Aviso getCalendarsByName:', eNome);
  }

  // 2. Busca em todos os calendários do usuário
  try {
    const todos = CalendarApp.getAllCalendars();
    
    // Correspondência direta com o nome Fabiane Satie
    for (let i = 0; i < todos.length; i++) {
      const c = todos[i];
      const nome = (c.getName() || '').trim().toLowerCase();
      if (nome === nomeAlvoLower || (nome.includes('fabiane') && nome.includes('satie'))) {
        console.log('Agenda Fabiane Satie localizada em getAllCalendars:', c.getName());
        return c;
      }
    }

    // Correspondência com nome contendo "Fabiane"
    for (let i = 0; i < todos.length; i++) {
      const c = todos[i];
      const nome = (c.getName() || '').trim().toLowerCase();
      if (nome.includes('fabiane')) {
        console.log('Agenda contendo Fabiane localizada:', c.getName());
        return c;
      }
    }

    // Busca pela agenda com checkbox ativo (isSelected() === true), ignorando listas auxiliares desmarcadas
    const ignorar = ['aniversários', 'aniversarios', 'tarefas', 'feriados', 'arena magalu'];
    for (let i = 0; i < todos.length; i++) {
      const c = todos[i];
      const nome = (c.getName() || '').trim().toLowerCase();
      const deveIgnorar = ignorar.some(ign => nome.includes(ign));
      if (deveIgnorar) continue;

      try {
        if (typeof c.isSelected === 'function' && c.isSelected()) {
          console.log('Agenda marcada na interface localizada:', c.getName());
          return c;
        }
      } catch (eSel2) {}
    }
  } catch (eAll) {
    console.warn('Aviso getAllCalendars:', eAll);
  }

  // 3. Fallback seguro para o calendário principal da conta
  try {
    const calPadrao = CalendarApp.getDefaultCalendar();
    if (calPadrao) {
      console.log('Utilizando calendário padrão:', calPadrao.getName());
      return calPadrao;
    }
  } catch (eDef) {
    console.warn('Aviso getDefaultCalendar:', eDef);
  }

  return null;
}

/**
 * Busca eventos EXCLUSIVAMENTE da agenda 'Fabiane Satie'
 * para o intervalo em visualização (Segunda a Domingo).
 * 
 * @param {string} dataInicioISO Data inicial YYYY-MM-DD
 * @param {string} dataFimISO Data final YYYY-MM-DD
 * @return {Array<Object>} Lista de eventos formatados para a UI
 */
function Service_Calendar_buscarEventosSemana(dataInicioISO, dataFimISO) {
  try {
    const eventosFormatados = [];
    const idsVistos = {};
    
    // Converte datas para instâncias de Date considerando início e fim do dia
    const dtInicio = Utils_parseDataSegura(dataInicioISO) || new Date();
    dtInicio.setHours(0, 0, 0, 0);
    
    const dtFim = Utils_parseDataSegura(dataFimISO) || new Date();
    dtFim.setHours(23, 59, 59, 999);
    
    // 1. Obtém unicamente a agenda Fabiane Satie
    const calAlvo = Service_Calendar_obterCalendarioAlvo();
    if (!calAlvo) {
      console.warn('Nenhuma agenda disponível encontrada.');
      return [];
    }

    console.log('Buscando eventos estritamente na agenda:', calAlvo.getName());
    
    // 2. Busca eventos apenas desse calendário
    try {
      const eventosBrutos = calAlvo.getEvents(dtInicio, dtFim);
      for (let i = 0; i < eventosBrutos.length; i++) {
        const ev = eventosBrutos[i];
        const id = ev.getId();
        if (idsVistos[id]) continue;
        idsVistos[id] = true;
        
        const isDiaInteiro = ev.isAllDayEvent();
        const inicio = ev.getStartTime();
        const fim = ev.getEndTime();
        const dataEventoISO = Utils_formatarDataISO(inicio);
        const desc = ev.getDescription() || '';
        const loc = ev.getLocation() || '';
        
        // Extrai link do Google Meet se presente na descrição ou local
        let linkMeet = '';
        const matchMeet = (desc + ' ' + loc).match(/https:\/\/meet\.google\.com\/[a-z0-9-]+/i);
        if (matchMeet) {
          linkMeet = matchMeet[0];
        }
        
        const horaInicioStr = isDiaInteiro ? 'Dia Inteiro' : Utils_formatarHora(inicio);
        const horaFimStr = isDiaInteiro ? '' : Utils_formatarHora(fim);
        
        eventosFormatados.push({
          id: id,
          titulo: ev.getTitle() || '(Sem Título)',
          dataISO: dataEventoISO,
          dataBR: Utils_formatarDataBR(dataEventoISO),
          horaInicio: horaInicioStr,
          horaFim: horaFimStr,
          inicio: inicio ? Utilities.formatDate(inicio, Utils_obterFusoHorario(), "yyyy-MM-dd'T'HH:mm:ss") : '',
          fim: fim ? Utilities.formatDate(fim, Utils_obterFusoHorario(), "yyyy-MM-dd'T'HH:mm:ss") : '',
          diaInteiro: isDiaInteiro,
          local: loc,
          descricao: desc.substring(0, 300),
          linkMeet: linkMeet,
          nomeCalendario: calAlvo.getName(),
          cor: ev.getColor() || '#4f634f'
        });
      }
    } catch (eCal) {
      console.warn('Erro ao obter eventos da agenda:', eCal);
    }
    
    // Ordena eventos cronologicamente
    eventosFormatados.sort((a, b) => {
      if (a.dataISO !== b.dataISO) {
        return a.dataISO.localeCompare(b.dataISO);
      }
      if (a.diaInteiro && !b.diaInteiro) return -1;
      if (!a.diaInteiro && b.diaInteiro) return 1;
      return (a.horaInicio || '').localeCompare(b.horaInicio || '');
    });
    
    return eventosFormatados;
  } catch (error) {
    console.error('Aviso: Falha controlada ao ler eventos do Google Calendar:', error.message);
    return [];
  }
}

/**
 * Cria compromisso na agenda 'Fabiane Satie'
 * 
 * @param {Object} eventoPayload Dados do evento a ser criado
 * @return {Object} Confirmação
 */
function Service_Calendar_criarEventoRapido(eventoPayload) {
  try {
    const cal = Service_Calendar_obterCalendarioAlvo();
    if (!cal) throw new Error('Calendário Fabiane Satie indisponível.');
    
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
      dataISO: dataISO,
      calendario: cal.getName()
    }, 'Evento adicionado ao Google Calendar na agenda ' + cal.getName() + '.');
  } catch (error) {
    return Utils_criarRespostaErro(error);
  }
}
