// =============================================================================
// GP360 ECOSYSTEM - GESTÃO DE PROJETOS E AGENDAS
// Arquivo: Config.gs
// Responsabilidade: Constantes globais, modelos de dados, paleta pastel e bootstrap
// =============================================================================

/**
 * Título oficial e versão do ecossistema GP360
 */
const APP_TITLE = 'GP360 - Gestão de Projetos e Agendas';
const APP_VERSION = '1.0.0';

/**
 * Chave de propriedade do script para persistência do ID da Planilha Google Sheets
 */
const SPREADSHEET_ID_PROPERTY = 'GP360_PROJETOS_SPREADSHEET_ID';
const DEFAULT_SPREADSHEET_NAME = 'GP360 - Gestão de Projetos e Agendas (DB)';

/**
 * Mapeamento canônico das abas do banco de dados relacional
 */
const SHEET_NAMES = Object.freeze({
  PROJETOS: 'PROJETOS',
  TAREFAS: 'TAREFAS',
  HISTORICO_TEMPO: 'HISTORICO_TEMPO',
  CONFIGURACOES: 'CONFIGURACOES'
});

/**
 * Definição estrita das colunas de cada aba da planilha (Ordem mandatória)
 */
const SHEET_HEADERS = Object.freeze({
  PROJETOS: [
    'ID_Projeto',
    'Nome_Projeto',
    'Area_Contexto',
    'Cor_Pastel',
    'Status',
    'Data_Criacao',
    'Observacoes'
  ],
  TAREFAS: [
    'ID_Tarefa',
    'ID_Projeto',
    'Titulo',
    'Proximos_Passos',
    'Data_Programada',
    'Prioridade',
    'Status',
    'Data_Inicio_Aguardando',
    'Motivo_Aguardando',
    'Tempo_Total_Segundos',
    'Data_Conclusao',
    'Criado_Em',
    'Atualizado_Em'
  ],
  HISTORICO_TEMPO: [
    'ID_Sessao',
    'ID_Tarefa',
    'Data_Execucao',
    'Hora_Inicio',
    'Hora_Fim',
    'Duracao_Segundos',
    'Usuario'
  ],
  CONFIGURACOES: [
    'Chave',
    'Valor'
  ]
});

/**
 * Status padrão de tarefas
 */
const STATUS_TAREFA = Object.freeze({
  PENDENTE: 'Pendente',
  EM_ANDAMENTO: 'Em Andamento',
  AGUARDANDO: 'Aguardando',
  CONCLUIDA: 'Concluída',
  ARQUIVADA: 'Arquivada'
});

/**
 * Status padrão de projetos
 */
const STATUS_PROJETO = Object.freeze({
  ATIVO: 'Ativo',
  PAUSADO: 'Pausado',
  CONCLUIDO: 'Concluído',
  ARQUIVADO: 'Arquivado'
});

/**
 * Níveis de prioridade das tarefas
 */
const PRIORIDADE_TAREFA = Object.freeze({
  ALTA: 'Alta',
  MEDIA: 'Média',
  BAIXA: 'Baixa'
});

/**
 * Paleta de cores pastéis moderna para personalização de projetos e identificadores visuais
 */
const PALETA_CORES_PASTEL = Object.freeze([
  { id: 'lavanda', nome: 'Lavanda Suave', bg: '#EDE9FE', texto: '#5B21B6', border: '#DDD6FE' },
  { id: 'indigo', nome: 'Índigo Suave', bg: '#E0E7FF', texto: '#3730A3', border: '#C7D2FE' },
  { id: 'azul', nome: 'Azul Céu Suave', bg: '#DBEAFE', texto: '#1E40AF', border: '#BFDBFE' },
  { id: 'menta', nome: 'Verde Menta Suave', bg: '#D1FAE5', texto: '#065F46', border: '#A7F3D0' },
  { id: 'ambar', nome: 'Âmbar Suave', bg: '#FEF3C7', texto: '#92400E', border: '#FDE68A' },
  { id: 'pessego', nome: 'Pêssego Suave', bg: '#FFEDD5', texto: '#9A3412', border: '#FED7AA' },
  { id: 'rosa', nome: 'Rosa Suave', bg: '#FFE4E6', texto: '#9F1239', border: '#FECDD3' },
  { id: 'purpura', nome: 'Púrpura Suave', bg: '#F3E8FF', texto: '#6B21A8', border: '#E9D5FF' },
  { id: 'ardosia', nome: 'Ardósia Suave', bg: '#F1F5F9', texto: '#334155', border: '#CBD5E1' }
]);

/**
 * Cores fundamentais do Design System Soft Pastel
 */
const CORES_SISTEMA = Object.freeze({
  PRIMARIA: '#7C3AED',
  PRIMARIA_BADGE: '#F3E8FF',
  SUCESSO: '#10B981',
  SUCESSO_BADGE: '#ECFDF5',
  ALERTA: '#F59E0B',
  ALERTA_BADGE: '#FFFBEB',
  CRITICO: '#F43F5E',
  CRITICO_BADGE: '#FFF1F2',
  FUNDO: '#F8FAFC',
  SUPERFICIE: '#FFFFFF',
  TEXTO_PRINCIPAL: '#0F172A',
  TEXTO_MUTED: '#64748B'
});

/**
 * Regras estritas de color coding e faixas de tempo decorrido para status 'Aguardando'
 */
const FAIXAS_AGUARDANDO = Object.freeze({
  FAIXA_1: {
    minDias: 1,
    maxDias: 2,
    bg: '#EFF6FF',
    texto: '#1E40AF',
    border: '#BFDBFE',
    descricao: '1 a 2 dias aguardando (Azul Suave)'
  },
  FAIXA_2: {
    minDias: 3,
    maxDias: 5,
    bg: '#FFFBEB',
    texto: '#92400E',
    border: '#FDE68A',
    descricao: '3 a 5 dias aguardando (Laranja/Pêssego Suave)'
  },
  FAIXA_3: {
    minDias: 6,
    maxDias: Infinity,
    bg: '#FFF1F2',
    texto: '#9F1239',
    border: '#FECDD3',
    descricao: '6+ dias aguardando (Rosa/Coral Suave)'
  }
});

/**
 * Função de entrada doGet() para renderização da Single Page Application (SPA).
 * Configura viewport móvel e segurança do Web App.
 * 
 * @param {Object} e Parâmetros da requisição HTTP GET
 * @return {HtmlOutput} Página HTML avaliada e configurada
 */
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('Index');
    
    template.appTitle = APP_TITLE;
    template.appVersion = APP_VERSION;
    template.usuarioLogado = Security_obterEmailUsuarioAtivo();
    
    return template.evaluate()
      .setTitle(APP_TITLE)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;color:#9F1239;background:#FFF1F2;border-radius:12px;">' +
      '<h2>Falha na inicialização do Web App GP360</h2>' +
      '<p>' + error.message + '</p>' +
      '</div>'
    ).setTitle('Erro - GP360');
  }
}

/**
 * Função auxiliar para modularização de arquivos HTML (CSS e JS) em templates GAS.
 * Permite sintaxe <?!= include('NomeDoArquivo'); ?> no Index.html.
 * 
 * @param {string} filename Nome do arquivo HTML (sem extensão .html)
 * @return {string} Conteúdo HTML puro
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
