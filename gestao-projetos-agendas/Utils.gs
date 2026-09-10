// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Utils.gs
// Responsabilidade: Formatadores de data/hora, cálculo de dias e funções utilitárias
// =============================================================================

/**
 * Regra global de higienização de filiais do ecossistema GP360 (Séries 3000/4000 para Série Original).
 * Padrão canônico obrigatório.
 * 
 * @param {string|number} val ID da filial bruto.
 * @return {string} ID da filial higienizado.
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Retorna o fuso horário padrão da aplicação (prioriza fuso do script ou São Paulo)
 * 
 * @return {string} Fuso horário IANA
 */
function Utils_obterFusoHorario() {
  try {
    return Session.getScriptTimeZone() || 'America/Sao_Paulo';
  } catch (e) {
    return 'America/Sao_Paulo';
  }
}

/**
 * Retorna a data de hoje no formato estrito YYYY-MM-DD considerando o fuso horário local
 * 
 * @return {string} Data no formato YYYY-MM-DD
 */
function Utils_obterDataHojeISO() {
  return Utilities.formatDate(new Date(), Utils_obterFusoHorario(), 'yyyy-MM-dd');
}

/**
 * Converte qualquer representação de data para o padrão estrito YYYY-MM-DD
 * 
 * @param {Date|string|number} dataValor Valor a ser convertido
 * @return {string} Data no formato YYYY-MM-DD ou string vazia
 */
function Utils_formatarDataISO(dataValor) {
  if (!dataValor) return "";
  try {
    if (dataValor instanceof Date) {
      if (isNaN(dataValor.getTime())) return "";
      return Utilities.formatDate(dataValor, Utils_obterFusoHorario(), 'yyyy-MM-dd');
    }
    
    let str = String(dataValor).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
      return str;
    }
    
    // Tratamento de formato brasileiro dd/MM/yyyy
    if (str.includes('/')) {
      const partes = str.split(' ')[0].split('/');
      if (partes.length === 3) {
        const dia = partes[0].padStart(2, '0');
        const mes = partes[1].padStart(2, '0');
        const ano = partes[2].length === 2 ? '20' + partes[2] : partes[2];
        return `${ano}-${mes}-${dia}`;
      }
    }
    
    const dt = new Date(dataValor);
    if (!isNaN(dt.getTime())) {
      return Utilities.formatDate(dt, Utils_obterFusoHorario(), 'yyyy-MM-dd');
    }
    return "";
  } catch (e) {
    return "";
  }
}

/**
 * Formata data para exibição brasileira amigável dd/MM/yyyy
 * 
 * @param {Date|string|number} dataValor Valor da data
 * @return {string} Data formatada ou string vazia
 */
function Utils_formatarDataBR(dataValor) {
  if (!dataValor) return "";
  try {
    const iso = Utils_formatarDataISO(dataValor);
    if (!iso) return "";
    const partes = iso.split('-');
    return `${partes[2]}/${partes[1]}/${partes[0]}`;
  } catch (e) {
    return "";
  }
}

/**
 * Formata hora no padrão brasileiro HH:mm
 * 
 * @param {Date|string} dataValor Objeto Date ou string de data/hora
 * @return {string} Hora no padrão HH:mm
 */
function Utils_formatarHora(dataValor) {
  if (!dataValor) return "";
  try {
    const dt = (dataValor instanceof Date) ? dataValor : new Date(dataValor);
    if (isNaN(dt.getTime())) return "";
    return Utilities.formatDate(dt, Utils_obterFusoHorario(), 'HH:mm');
  } catch (e) {
    return "";
  }
}

/**
 * Faz parse seguro de string para objeto Date ajustado para a meia-noite local
 * 
 * @param {string|Date} dataValor String YYYY-MM-DD ou Date
 * @return {Date|null} Objeto Date ou null se inválido
 */
function Utils_parseDataSegura(dataValor) {
  if (!dataValor) return null;
  if (dataValor instanceof Date) {
    return isNaN(dataValor.getTime()) ? null : dataValor;
  }
  
  const iso = Utils_formatarDataISO(dataValor);
  if (!iso) return null;
  
  const partes = iso.split('-');
  const ano = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1;
  const dia = parseInt(partes[2], 10);
  
  const dt = new Date(ano, mes, dia, 0, 0, 0, 0);
  return isNaN(dt.getTime()) ? null : dt;
}

/**
 * Soma ou subtrai dias a partir de uma data no formato YYYY-MM-DD
 * 
 * @param {string} dataISO Data inicial no formato YYYY-MM-DD
 * @param {number} dias Quantidade de dias a adicionar (positivo ou negativo)
 * @return {string} Nova data no formato YYYY-MM-DD
 */
function Utils_adicionarDias(dataISO, dias) {
  const dt = Utils_parseDataSegura(dataISO) || new Date();
  dt.setDate(dt.getDate() + dias);
  return Utils_formatarDataISO(dt);
}

/**
 * Calcula o número de dias corridos decorridos entre duas datas
 * 
 * @param {string|Date} dataInicio Data de início
 * @param {string|Date} dataFim Data final (se omitida, usa hoje)
 * @return {number} Dias decorridos inteiros (>= 0)
 */
function Utils_calcularDiasDecorridos(dataInicio, dataFim) {
  if (!dataInicio) return 0;
  
  const dtInicio = Utils_parseDataSegura(dataInicio);
  const dtFim = dataFim ? Utils_parseDataSegura(dataFim) : Utils_parseDataSegura(Utils_obterDataHojeISO());
  
  if (!dtInicio || !dtFim) return 0;
  
  const diffMilissegundos = dtFim.getTime() - dtInicio.getTime();
  const diffDias = Math.floor(diffMilissegundos / (1000 * 60 * 60 * 24));
  return diffDias >= 0 ? diffDias : 0;
}

/**
 * Converte segundos acumulados no formato legível HH:MM:SS
 * 
 * @param {number} totalSegundos Quantidade de segundos
 * @return {string} String no formato 00:00:00
 */
function Utils_formatarSegundosParaHHMMSS(totalSegundos) {
  const segs = Math.max(0, parseInt(totalSegundos, 10) || 0);
  const horas = Math.floor(segs / 3600);
  const minutos = Math.floor((segs % 3600) / 60);
  const segundos = segs % 60;
  
  const hh = String(horas).padStart(2, '0');
  const mm = String(minutos).padStart(2, '0');
  const ss = String(segundos).padStart(2, '0');
  
  return `${hh}:${mm}:${ss}`;
}

/**
 * Calcula o intervalo de dias da semana (Segunda a Domingo) contendo a data de referência
 * 
 * @param {string} dataReferenciaISO Data YYYY-MM-DD ou nula (assume hoje)
 * @return {Object} Metadados da semana e array dos 7 dias formatados
 */
function Utils_obterIntervaloSemana(dataReferenciaISO) {
  const dataBase = dataReferenciaISO ? Utils_parseDataSegura(dataReferenciaISO) : Utils_parseDataSegura(Utils_obterDataHojeISO());
  const dt = dataBase || new Date();
  
  // No JS: 0 = Domingo, 1 = Segunda, ..., 6 = Sábado
  const diaSemanaIndex = dt.getDay();
  // Diferença até a Segunda-feira: se Domingo (0), recua 6 dias. Senão, recua (diaSemanaIndex - 1)
  const diffParaSegunda = (diaSemanaIndex === 0) ? -6 : (1 - diaSemanaIndex);
  
  const segunda = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate() + diffParaSegunda, 0, 0, 0, 0);
  const hojeISO = Utils_obterDataHojeISO();
  
  const nomesDias = [
    { curto: 'SEG', completo: 'Segunda-feira' },
    { curto: 'TER', completo: 'Terça-feira' },
    { curto: 'QUA', completo: 'Quarta-feira' },
    { curto: 'QUI', completo: 'Quinta-feira' },
    { curto: 'SEX', completo: 'Sexta-feira' },
    { curto: 'SÁB', completo: 'Sábado' },
    { curto: 'DOM', completo: 'Domingo' }
  ];
  
  const nomesMeses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  
  const diasSemana = [];
  for (let i = 0; i < 7; i++) {
    const diaAtual = new Date(segunda.getFullYear(), segunda.getMonth(), segunda.getDate() + i, 0, 0, 0, 0);
    const dataISO = Utils_formatarDataISO(diaAtual);
    
    diasSemana.push({
      dataISO: dataISO,
      dataBR: Utils_formatarDataBR(diaAtual),
      diaSemanaCurto: nomesDias[i].curto,
      diaSemanaCompleto: nomesDias[i].completo,
      numeroDia: diaAtual.getDate(),
      mesCurto: nomesMeses[diaAtual.getMonth()],
      ano: diaAtual.getFullYear(),
      isHoje: (dataISO === hojeISO)
    });
  }
  
  const domingo = new Date(segunda.getFullYear(), segunda.getMonth(), segunda.getDate() + 6, 23, 59, 59, 999);
  
  return {
    segundaISO: diasSemana[0].dataISO,
    domingoISO: diasSemana[6].dataISO,
    dataReferenciaISO: Utils_formatarDataISO(dt),
    hojeISO: hojeISO,
    diasSemana: diasSemana
  };
}

/**
 * Gera um ID único com prefixo legível e tolerante a colisões
 * 
 * @param {string} prefixo Ex: 'PRJ', 'TSK', 'SES'
 * @return {string} Identificador único
 */
function Utils_gerarId(prefixo) {
  const pre = prefixo ? String(prefixo).toUpperCase().trim() : 'ID';
  const timestamp = Date.now().toString(36).toUpperCase();
  const aleatorio = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${pre}-${timestamp}-${aleatorio}`;
}

/**
 * Cria envelope padrão de resposta de sucesso para as chamadas RPC
 * 
 * @param {*} dados Payload da resposta
 * @param {string} mensagem Mensagem informativa
 * @return {Object} Objeto estruturado
 */
function Utils_criarRespostaSucesso(dados, mensagem) {
  return {
    sucesso: true,
    dados: dados,
    mensagem: mensagem || 'Operação realizada com sucesso.',
    timestamp: new Date().toISOString()
  };
}

/**
 * Cria envelope padrão de resposta de erro para as chamadas RPC
 * 
 * @param {*} erro Mensagem ou instância de Error
 * @param {string} codigo Código de erro de negócio
 * @return {Object} Objeto estruturado
 */
function Utils_criarRespostaErro(erro, codigo) {
  const msg = (erro instanceof Error) ? erro.message : String(erro || 'Ocorreu um erro inesperado.');
  return {
    sucesso: false,
    erro: msg,
    codigo: codigo || 'ERRO_INTERNO',
    timestamp: new Date().toISOString()
  };
}
