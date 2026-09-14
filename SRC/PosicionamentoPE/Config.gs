// =====================================================================
// CONFIGURAÇÕES GLOBAIS E ROTEADOR (MÓDULO: EVENTO REGIONAL PILOTO)
// =====================================================================

// Planilha Base (Consulta de Ativos e Lojas - DADOS_ATIVOS / DADOS_LOJAS)
const SPREADSHEET_BASE_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; 

// Planilha de Respostas (Nova base criada - Detalhamento Regional)
const SPREADSHEET_DESTINO_ID = '1K9qtR28oeacW9U1prx24NhqOz_p89dabY0rn3mBHIV8'; 

// E-mails Oficiais
const EMAIL_ROSA = 'rosa.santos@magazineluiza.com.br';
const EMAIL_AGENCIA = 'posicionamentomagalu@ltnbrasil.com.br';

// Data limite para o preenchimento (Trava Temporal)
const DATA_LIMITE_PREENCHIMENTO = new Date('2026-09-21T23:59:59');

// Super Admins: Usuários que ignoram bloqueios de regional, férias e DADOS_USUARIOS
const SUPER_ADMINS = ['fabiane.satie@magazineluiza.com.br', 'rosa.santos@magazineluiza.com.br'];

/**
 * Função principal que atua como Roteador do Web App.
 * Identifica se a requisição é para o Formulário ou Dashboard.
 */
function doGet(e) {
  var page = e.parameter.page;
  
  if (page === 'dashboard') {
    // Carrega a visão do Coordenador (Dashboard Bento Grid)
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Dashboard Coordenadores - Evento Regional')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    // Carrega a visão do Gerente (Formulário)
    // Validação do Pilar de Temporização
    var dataAtual = new Date();
    if (dataAtual > DATA_LIMITE_PREENCHIMENTO) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:sans-serif; padding:50px; text-align:center; color:#9F1239; background:#FFF1F2; border-radius:10px; border: 1px solid #F43F5E;">' +
        '<h2>Período Encerrado</h2><p>O prazo para preenchimento de dados do Evento Regional foi finalizado em 21/09/2026.</p>' +
        '<p>Em caso de dúvidas, contate a gestora responsável.</p></div>'
      );
    }
    
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Formulário de Coleta - Evento Regional')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}
