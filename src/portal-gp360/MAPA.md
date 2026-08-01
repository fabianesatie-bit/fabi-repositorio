MAPA DA ARQUITETURA DO MÓDULO: PORTAL GP 360
Subfolder no GitHub: src/portal-gp360/
1. Mapeamento de Constantes e Variáveis Globais
IDs de Planilhas / Bancos de Dados:
SPREADSHEET_ID: '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA' (DB_MASTER)
SPREADSHEET_DASH: '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc' (Base de Indicadores Geral)
SPREADSHEET_LOG_ID: '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY' (Planilha Auditoria Exclusiva)
IDs de Pastas do Google Drive:
EVIDENCIAS_FOLDER_ID: '1v28G-ZDd6yQpjTUvBNcODlpxkM5AeaQZ' (Pasta do Drive para Evidências)
Listas de Acesso / Whitelist:
SUPER_ADMINS_EMAILS:
"fabiane.satie@magazineluiza.com.br"
"gplojas@magazineluiza.com.br"
"tarcisio.maniglia@magazineluiza.com.br"
Outras Configurações Globais (Client-side em Scripts_Core.html):
dadosGlobais: Armazena o estado global da aplicação em memória (usuário, moedas, histórico, etc).
arquivoColadoMemoria: Armazena payload em base64 da imagem oriunda de Ctrl+V.
fileIdLightboxAtivo: Mantém o ID da imagem em exibição no modal de galeria.
urlsRedirecionamento: Dicionário contendo as URLs fixas de painéis externos ("Relatorios/Feedback/Acompanhamento" e "Atendimento Social").
motivosRedirecionamento: Chaves do dicionário de URLs extraídas em array.
2. Mapeamento Detalhado de Funções por Arquivo
Config.gs
doGet(): Constrói a interface a partir de Index.html, permitindo inclusão em iframe via ALLOWALL e injetando a tag de viewport para responsividade.
include(filename): Função auxiliar para ler e injetar o conteúdo de arquivos HTML dentro de outros.
Security.gs
registrarAuditoria(evento, detalhe): Registra log silencioso na planilha AUDITORIA, capturando o e-mail ativo e timestamp. Trata erros silenciosamente (try/catch).
obterControleAcesso(email): Constrói o perfil do usuário consultando a aba DADOS_USUARIOS. Mapeia privilégios booleanos (isSuperAdmin, isAdmin, isConfigAdmin, isGerenteGP) baseando-se em listas de e-mail e strings de cargo/acesso.
validarAcessoFilial(email, filialId): Verifica em DADOS_LOJAS se o e-mail possui jurisdição geográfica (Regional) correspondente à loja consultada.
Utils.gs
extrairIdDrive(linkOuId): Extrai apenas o hash ID alfanumérico (25+ caracteres) de qualquer URL completa do Google Drive via Regex.
formatarDataSegura(dataValor): Padroniza outputs de dados temporais para string dd/MM/yyyy, suportando instâncias Date e manipulações de string crua.
obterDataRawSegura(dataValor): Converte strings de datas variadas (incluindo padrões invertidos de tela e backend) para timestamps (getTime()) para algoritmos de ordenação e filtros.
Service_Data.gs
obterDadosIniciais(): Orquestrador massivo. Compila perfil, lista de lojas filtradas por RLS, parseia histórico de DADOS_LANCAMENTOS validando visibilidade cruzada (Overlap de Regionais), calcula placares de gamificação (moedas totais, moedas do mês, badges, fase da montanha), gastos mensais, extrai recados do mural e compila dropdowns de naturezas.
carregarNaturezasSeguras(ss): Lê aba CONFIGURAÇÕES, removendo itens nulos e ignorando chaves restritas/fixas (array fixas hardcoded).
buscarIndicadoresLoja(filialId): Usa CacheService ('IND_LOJA_' + id) por 300 segundos. Verifica validarAcessoFilial e varre planilhas de SPREADSHEET_DASH para obter Vendas, NPS, Banco de Horas, Quadro e Tx de Desligamento, aplicando tratamento numérico pesado linha a linha.
Service_Write.gs
registrarAtividade(dados, fileData): Aplica LockService. Checa privilégios, destrói cache da filial afetada (CacheService.remove), decoda Base64 da evidência salvando no Drive e executa appendRow exato em 27 colunas do DADOS_LANCAMENTOS.
excluirLancamento(idLancamento): Aplica LockService. Checa permissão (Dono do registro ou SuperAdmin), varre array em busca do ID da linha e aplica deleteRow(). Limpa o cache de indicadores da filial deletada.
Service_Admin.gs
getDicionarioPremios(): Recupera CRUD de prêmios da aba DICIONARIO_PREMIOS.
salvarPremio(premio): Insere (appendRow) ou atualiza (setValue) prêmio. Restrito a GerenteGP ou isSuperAdmin.
deletarPremio(idPremio): Remove a linha do prêmio na aba de dicionário.
adicionarNatureza(novaNatureza): Insere linha em CONFIGURAÇÕES. Restrito a Administradores. Trava inserções duplicadas.
excluirNatureza(naturezaTexto): Varre CONFIGURAÇÕES e remove a linha da natureza especificada.
publicarAviso(mensagem, diretoriaAlvo): Insere em DADOS_AVISOS. Filtra DADOS_USUARIOS para enviar notificação em lote via MailApp.sendEmail usando cópia oculta (BCC).
excluirAvisoPlanilha(dataAviso, mensagemAviso): Remove registro da aba DADOS_AVISOS validando string de data e texto.
Service_Drive.gs
getEvidenciasUsuario(): Recupera as últimas 150 evidências visuais de DADOS_LANCAMENTOS. Verifica permissões e monta objeto complexo gerando URL de miniatura (thumbnail?id=) e URL de download.
getImagemBase64(fileId): Converte um binário de imagem do Drive num Base64 payload, burlando o bloqueio de CORS nativo do canvas do navegador.
Index.html
Estrutura Visual / Divs Base:
#bloqueio-tela: Overlay de erro para usuários sem permissão.
#loading-splash: Tela de loading e sincronização inicial.
#app-container: Grid principal (Sidebar esquerdo + Área de Conteúdo).
Abas Principais (.view-aba):
#aba-home: Dashboards, métricas, mural de avisos, ranking e trilha da montanha.
#aba-lancamento: Formulário extenso e dinâmico para registrar viagens, visitas e roteiros.
#aba-gestao: Timeline vertical de atividades (view crua e admin de remoções).
#aba-galeria: Grid de imagens (thumbnails) puxadas do Google Drive.
#aba-historico: Timeline de despesas operacionais com botão de impressão.
#aba-consultas: Grid de atalhos e iframe injetável para acesso de sistemas externos.
#aba-configuracoes: (Painel de Administração) CRUD de prêmios e naturezas.
Modais:
#toast: Notificações rápidas no canto inferior.
#modal-lightbox: Visualizador de imagens em zoom da Galeria de Evidências.
#modal-novo-aviso: Formulário de criação de recado para a Home.
#modal-km-avulso: Formulário isolado para registro de KM Administrativo sem checklist de loja.
Scripts_Core.html
window.onload: Inicialização, set de data no formulário e chamada assíncrona de obterDadosIniciais(). Injeta mocks em caso de debug local fora do Apps Script.
salvarAtividade(): Coleta valores do DOM, define lógicas condicionais (se roteiro redirect, desvia), monta DTO de inserção (com verificação de arquivo de imagem vs payload Base64) e despacha ao backend.
chamarBackendSalvar(dados, fileData, btn): Executa API externa, gerencia Loader no botão, notifica toast, atira confetes de sucesso, reseta formulário e atua no Resync de dados.
deletarHistorico(id): Invoca frontend confirmation, injeta opacidade fantasma na linha, faz drop da linha do array local e solicita deleção server-side.
salvarKMAvulso(): Cria DTO específico para KM isolado e manda via registrarAtividade.
salvarNovoAviso(), deletarAviso(data, msg): Manipuladores CRUD para os recados corporativos.
addNatureza(), delNatureza(texto): Manipuladores CRUD da tela de Admin para ditar dropdown.
salvarPremioAdmin(), deletarPremioAdmin(id): Manipuladores CRUD do dicionário de moedas/gamificação.
Scripts_UI.html
limparOutrasEvidencias(origem): UX Handler para forçar Single Choice entre Upload Local, Paste Clipboard ou Drive Link.
customConfirm(message, onConfirm): Substituto customizado modal do window.confirm.
mostrarToast(msg): Gatilho CSS de notificação.
dispararConfetes(): Gatilho visual em canvas-confetti.
usarIniciaisComoFallback(), usarIniciais(nome): Fallback de imagem de perfil baseada na inicial do nome.
toggleMenu(), mudarAba(abaId, element): Controladores de Viewport e abas.
abrirPainelExterno(), fecharPainelExterno(): Manipula Src de iframe-externo.
renderizarGamificacaoVisual(res): Modifica DOM com a meta de fase na montanha e selos de badges.
renderizarRanking(ranking): Desenha tabela HTML listando apenas os 5 maiores em score.
renderizarAvisos(avisos), renderizarListaAdmin(naturezas), renderizarPremiosAdmin(premios): DOM Injection Functions para grids corporativas.
toggleTodasAbrangencias(checkbox): Select-all macro de formulário.
alternarEscopo(valor), atualizarSelectNaturezas(), verificarCamposDinamicos(): Máquina de estados UI do formulário de lançamento, escondendo/mostrando campos baseado no escopo e motivo.
calcularDespesas(), calcularDespesasAvulso(): Multiplica Rate KM * Distância e soma demais variáveis de float, manipulando span de custo na tela.
mostrarIndicadoresLancamento(filialId): Faz o fetch assíncrono pro backend sobre a loja específica escolhida no form e desenha os cards do dashboard interativo do forms.
parseDataBR(dataString), limparFiltroDatas(), renderizarHistoricos(): Rotinas pesadas de parse local e reconstrução de lista vertical ul/li baseada em inputs date (min-max).
imprimirRelatorioKM(): Constrói uma nova janela popup window.open com HTML table de despesas filtradas e aplica win.print().
carregarGaleria(), abrirLightbox(ev), fecharLightbox(): Construtores de layout flex de cartões e acionadores do modal de view de fotos.
copiarImagemLightbox(): Interpela o server solicitando bytes raw, injeta num BLOB em memória e joga o objeto via API experimental do Browser (navigator.clipboard.write).
abrirModalAviso(), fecharModalAviso(), abrirModalAvulso(): Controladores CSS de modais.
3. Integrações, Leitura e Gravação no Google Sheets
Abas do Google Sheets Manipuladas (DB_MASTER):
DADOS_USUARIOS (Leitura)
DADOS_LOJAS (Leitura)
DADOS_LANCAMENTOS (Leitura por lotes e Gravação com appendRow e deleteRow)
DADOS_AVISOS (Leitura e Gravação com appendRow e deleteRow)
CONFIGURAÇÕES (Leitura e Gravação com appendRow e deleteRow)
DICIONARIO_PREMIOS (Leitura e Gravação/Edição com setValue)
Abas do Google Sheets Manipuladas (DB_DASH):
HISTORICO_VENDA_ANO, HISTORICO_NPS, HISTORICO_BH_ACUMULADO, HISTORICO_QUADRO, HISTORICO_TXDESL (Todas lidas massivamente em bloco bidimensional via getValues() sem gravações).
Abas do Google Sheets Manipuladas (DB_LOG):
AUDITORIA (Gravação exclusiva em loop sequencial via appendRow).
4. Trava de Segurança e Validações de Negócio
LockService:
Funções Críticas de Gravação (registrarAtividade, excluirLancamento) utilizam LockService.getScriptLock() com limite de 15 segundos prevendo concorrência pesada (Race Conditions) ao final de mês com a equipe inserindo KMs de viagens em massa.
Sanitização Base RLS (Row Level Security):
No Backend, a função validarAcessoFilial cruza a Regional que a matriz define em DADOS_USUARIOS para aquele e-mail logado, comparando-a com a Regional extraída em DADOS_LOJAS para impedir "espiada" indevida de dados sigilosos por diretores não alinhados à loja.
Aplica RLS na busca de indicadores no dashboard de formulário (buscarIndicadoresLoja), garantindo recusa via barreira técnica e não via bloqueio frontend.
Validação Híbrida de Administrador:
Backend recusa requisições puras via console de rede ao checar via obterControleAcesso os escopos: isSuperAdmin, isConfigAdmin e isGerenteGP. Apenas tais flags podem disparar e-mails globais (Avisos) ou gravar Dicionários Financeiros (Gamificação).
Cache Eficiente:
Inserido motor de cache local (CacheService) com 300 segundos de cooldown nas buscas de indicadores de Loja, evitando excesso de cotas da SpreadsheetApp da Google por re-seleções dos mesmos combos da filial em tela, com limpeza do cache imediata em registrarAtividade ou excluirLancamento.
