// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Security.gs
// Responsabilidade: Controle de concorrência com LockService e sanitização
// =============================================================================

/**
 * Executa uma operação crítica de banco de dados com bloqueio de concorrência distribuído.
 * Garante atomicidade em operações de leitura-modificação-escrita nas planilhas.
 * 
 * @param {Function} operacao Função de callback contendo a lógica a ser executada
 * @param {number} timeoutMs Tempo máximo de espera pelo lock em milissegundos (padrão: 10000)
 * @return {*} Resultado retornado pela função de callback
 * @throws {Error} Se o lock não puder ser adquirido dentro do prazo
 */
function Security_executarComLock(operacao, timeoutMs) {
  const tempoEspera = (typeof timeoutMs === 'number' && timeoutMs > 0) ? timeoutMs : 10000;
  const lock = LockService.getScriptLock();
  
  try {
    const obteveLock = lock.tryLock(tempoEspera);
    if (!obteveLock) {
      throw new Error('Servidor ocupado processando outra requisição concorrente. Por favor, aguarde alguns segundos e tente novamente.');
    }
    
    // Executa a operação protegida
    return operacao();
  } catch (error) {
    console.error('Erro na execução de bloco com Lock:', error);
    throw error;
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // Falha silenciosa no release caso o lock já tenha sido liberado
    }
  }
}

/**
 * Sanitiza valores textuais para prevenção estrita contra:
 * 1. Injeção de Fórmulas no Google Sheets (prevenção de "=", "+", "-", "@" no início da célula)
 * 2. Injeção de código script / XSS em visualizações HTML
 * 
 * @param {*} valor Dado de entrada a ser sanitizado
 * @return {*} Dado devidamente higienizado
 */
function Security_sanitizarString(valor) {
  if (valor === null || valor === undefined) return '';
  if (typeof valor !== 'string') return valor;
  
  let texto = valor.trim();
  
  // Prevenção contra injeção de fórmulas maliciosas nas planilhas
  if (/^[\=\+\-\@\t\r]/.test(texto)) {
    texto = "'" + texto;
  }
  
  // Higienização básica de tags HTML potencialmente perigosas
  texto = texto
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
    .replace(/javascript\s*:/gi, '');
    
  return texto;
}

/**
 * Higieniza recursivamente todas as chaves e valores de um objeto payload
 * 
 * @param {Object} obj Objeto recebido do cliente
 * @return {Object} Objeto higienizado
 */
function Security_sanitizarObjeto(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  
  if (Array.isArray(obj)) {
    return obj.map(item => Security_sanitizarObjeto(item));
  }
  
  const objetoLimpo = {};
  for (const chave in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, chave)) {
      const valor = obj[chave];
      if (typeof valor === 'string') {
        objetoLimpo[chave] = Security_sanitizarString(valor);
      } else if (typeof valor === 'object' && valor !== null) {
        objetoLimpo[chave] = Security_sanitizarObjeto(valor);
      } else {
        objetoLimpo[chave] = valor;
      }
    }
  }
  return objetoLimpo;
}

/**
 * Valida e normaliza o payload de uma tarefa antes da gravação no Google Sheets
 * 
 * @param {Object} tarefa Dados brutos da tarefa
 * @return {Object} Dados validados e estruturados
 */
function Security_validarPayloadTarefa(tarefa) {
  if (!tarefa || typeof tarefa !== 'object') {
    throw new Error('Payload da tarefa inválido ou não fornecido.');
  }
  
  const payloadLimpo = Security_sanitizarObjeto(tarefa);
  
  const titulo = String(payloadLimpo.Titulo || '').trim();
  if (!titulo) {
    throw new Error('O título da tarefa é obrigatório.');
  }
  
  // Normalização de prioridade
  let prioridade = String(payloadLimpo.Prioridade || PRIORIDADE_TAREFA.MEDIA).trim();
  const prioridadesValidas = Object.values(PRIORIDADE_TAREFA);
  if (!prioridadesValidas.includes(prioridade)) {
    prioridade = PRIORIDADE_TAREFA.MEDIA;
  }
  
  // Normalização de status
  let status = String(payloadLimpo.Status || STATUS_TAREFA.PENDENTE).trim();
  const statusValidos = Object.values(STATUS_TAREFA);
  if (!statusValidos.includes(status)) {
    status = STATUS_TAREFA.PENDENTE;
  }
  
  return {
    ID_Tarefa: payloadLimpo.ID_Tarefa ? String(payloadLimpo.ID_Tarefa).trim() : '',
    ID_Projeto: payloadLimpo.ID_Projeto ? String(payloadLimpo.ID_Projeto).trim() : '',
    Titulo: titulo,
    Proximos_Passos: String(payloadLimpo.Proximos_Passos || '').trim(),
    Data_Programada: String(payloadLimpo.Data_Programada || '').trim(),
    Prioridade: prioridade,
    Status: status,
    Data_Inicio_Aguardando: String(payloadLimpo.Data_Inicio_Aguardando || '').trim(),
    Motivo_Aguardando: String(payloadLimpo.Motivo_Aguardando || '').trim(),
    Tempo_Total_Segundos: parseInt(payloadLimpo.Tempo_Total_Segundos, 10) || 0
  };
}

/**
 * Valida e normaliza o payload de um projeto antes da gravação no Google Sheets
 * 
 * @param {Object} projeto Dados brutos do projeto
 * @return {Object} Dados validados e estruturados
 */
function Security_validarPayloadProjeto(projeto) {
  if (!projeto || typeof projeto !== 'object') {
    throw new Error('Payload do projeto inválido ou não fornecido.');
  }
  
  const payloadLimpo = Security_sanitizarObjeto(projeto);
  
  const nomeProjeto = String(payloadLimpo.Nome_Projeto || '').trim();
  if (!nomeProjeto) {
    throw new Error('O nome do projeto é obrigatório.');
  }
  
  // Cor pastel padrão (lavanda) se não informada
  let corPastel = String(payloadLimpo.Cor_Pastel || '#EDE9FE').trim();
  if (!/^#[0-9A-Fa-f]{6}$/.test(corPastel)) {
    corPastel = '#EDE9FE';
  }
  
  let status = String(payloadLimpo.Status || STATUS_PROJETO.ATIVO).trim();
  const statusValidos = Object.values(STATUS_PROJETO);
  if (!statusValidos.includes(status)) {
    status = STATUS_PROJETO.ATIVO;
  }
  
  return {
    ID_Projeto: payloadLimpo.ID_Projeto ? String(payloadLimpo.ID_Projeto).trim() : '',
    Nome_Projeto: nomeProjeto,
    Area_Contexto: String(payloadLimpo.Area_Contexto || '').trim(),
    Cor_Pastel: corPastel,
    Status: status,
    Observacoes: String(payloadLimpo.Observacoes || '').trim()
  };
}

/**
 * Identifica o e-mail do usuário ativo no Google Workspace
 * 
 * @return {string} E-mail do usuário autenticado
 */
function Security_obterEmailUsuarioAtivo() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email && email.trim() !== '') {
      return email.toLowerCase().trim();
    }
  } catch (e) {
    // Falha silenciosa em contextos restritos
  }
  
  try {
    const emailEfetivo = Session.getEffectiveUser().getEmail();
    if (emailEfetivo && emailEfetivo.trim() !== '') {
      return emailEfetivo.toLowerCase().trim();
    }
  } catch (e) {}
  
  return 'usuario.workspace@corporativo.com';
}

