/**
 * Cria o menu customizado na interface da planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automação Cursos')
    .addItem('Enviar Cobranças por Regional', 'abrirModalRegional')
    .addSeparator()
    .addItem('Instalar Gatilho Automático (Base)', 'instalarGatilhoAtualizacaoBase')
    .addToUi();
}

/**
 * Cria o gatilho (Trigger) programaticamente para detectar alterações estruturais na planilha.
 * Só deve ser executado uma vez pelo administrador da planilha.
 */
function instalarGatilhoAtualizacaoBase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Limpa gatilhos anteriores para evitar duplicidade
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'aoAlterarPlanilha') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('aoAlterarPlanilha')
    .forSpreadsheet(ss)
    .onChange();
    
  mostrarAlertaHTML("Gatilho Instalado", "O sistema agora monitorará automaticamente as atualizações na aba <strong>Base</strong>.<br><br>Quando atualizada, enviará cobranças para a diretoria <strong>MG/CO</strong> limitando a 1 disparo consolidado por dia para evitar SPAM.");
}

/**
 * Função executada automaticamente pelo gatilho 'onChange'.
 * Possui mecanismo de trava para evitar disparos em massa em atualizações tipo 'row-by-row'.
 */
function aoAlterarPlanilha(e) {
  // Como onChange não fornece o range exato da mesma forma que onEdit, validamos pela aba ativa no momento da alteração.
  const sheet = SpreadsheetApp.getActiveSheet();
  
  if (sheet.getName() !== "Base") {
    return; // Ignora se a edição não foi na aba Base
  }

  // LockService evita que edições simultâneas (ex: RPA ou colar vários dados) encavalem execuções.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    const properties = PropertiesService.getScriptProperties();
    const lastRun = properties.getProperty('LAST_RUN_BASE_MGCO');
    const today = new Date().toDateString();

    // Debounce: Se já rodou hoje, bloqueia novos envios para não lotar o e-mail dos gerentes
    if (lastRun === today) {
      return; 
    }

    // Processa envios focando apenas na diretoria MG/CO
    processarPendencias(null, "MG/CO");
    
    // Registra a data do envio para travar até amanhã
    properties.setProperty('LAST_RUN_BASE_MGCO', today);
  } catch (error) {
    console.error("Erro no gatilho automático: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exibe um alerta customizado em HTML para evitar bloqueios na UI.
 */
function mostrarAlertaHTML(titulo, mensagem) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; color: #333; line-height: 1.5; font-size: 14px; }
        button { margin-top: 20px; padding: 10px 20px; background-color: #0288d1; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; }
        button:hover { background-color: #0277bd; }
        .content { max-height: 200px; overflow-y: auto; }
      </style>
    </head>
    <body>
      <div class="content" id="mensagem">${mensagem.replace(/\n/g, '<br>')}</div>
      <div style="text-align: right;">
        <button onclick="google.script.host.close()">Fechar</button>
      </div>
    </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(450).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, titulo);
}

/**
 * Exibe um modal (HTML) com lista suspensa das Regionais alimentada dinamicamente.
 */
function abrirModalRegional() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDadosLoja = ss.getSheetByName("DADOS_LOJAS");
  
  if (!abaDadosLoja) {
    mostrarAlertaHTML("Erro Crítico", "Erro: Aba DADOS_LOJAS não encontrada.");
    return;
  }

  // Lê os dados pulando o cabeçalho. A Regional está na Coluna C (índice 2)
  const dados = abaDadosLoja.getDataRange().getValues().slice(1);
  const regionais = new Set();
  
  dados.forEach(row => {
    const reg = String(row[2]).trim().toUpperCase();
    if (reg) regionais.add(reg);
  });

  // Ordena alfabeticamente
  const regionaisArray = Array.from(regionais).sort();

  let optionsHtml = '';
  regionaisArray.forEach(reg => {
    // Escapa aspas para evitar quebra do HTML
    const safeReg = reg.replace(/"/g, '&quot;');
    optionsHtml += `<option value="${safeReg}">${safeReg}</option>`;
  });

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; color: #333; }
        select { width: 100%; padding: 10px; margin-bottom: 20px; border-radius: 5px; border: 1px solid #ccc; font-size: 14px; }
        button { padding: 10px 15px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 14px; }
        .btn-ok { background-color: #4CAF50; color: white; }
        .btn-cancel { background-color: #f44336; color: white; margin-left: 10px; }
        .btn-ok:hover { background-color: #45a049; }
        .btn-cancel:hover { background-color: #e53935; }
        .step-container { display: flex; flex-direction: column; }
        .actions { display: flex; justify-content: flex-end; }
      </style>
    </head>
    <body>
      <div id="step1" class="step-container">
        <label for="regional" style="margin-bottom: 10px; font-weight: bold;">Selecione a Regional:</label>
        <select id="regional">
          <option value="">-- Escolha uma Regional --</option>
          ${optionsHtml}
        </select>
        <div class="actions">
          <button class="btn-cancel" onclick="google.script.host.close()">Cancelar</button>
          <button class="btn-ok" onclick="pedirConfirmacao()" style="margin-left: 10px;">Avançar</button>
        </div>
      </div>

      <div id="step2" class="step-container" style="display:none;">
        <p style="margin-bottom: 20px;">Deseja processar e <strong>ENVIAR E-MAILS</strong> para todas as filiais da regional: <br><br><strong id="regName" style="color: #d32f2f; font-size: 16px;"></strong>?</p>
        <div class="actions">
          <button class="btn-cancel" onclick="voltar()">Voltar</button>
          <button class="btn-ok" onclick="enviar()" style="margin-left: 10px;">Confirmar Envio</button>
        </div>
        <p id="loading" style="display:none; color: #0288d1; margin-top:15px; font-weight: bold; text-align: center;">Processando envios... O alerta de relatório aparecerá em breve.</p>
      </div>

      <script>
        // Uso de manipulação segura do DOM evitando vulnerabilidades de injeção direta no innerHTML com dados de usuário
        function pedirConfirmacao() {
          const selectEl = document.getElementById('regional');
          const reg = selectEl.value;
          if(!reg) return; 
          
          document.getElementById('regName').textContent = reg;
          document.getElementById('step1').style.display = 'none';
          document.getElementById('step2').style.display = 'block';
        }
        
        function voltar() {
          document.getElementById('step2').style.display = 'none';
          document.getElementById('step1').style.display = 'block';
        }
        
        function enviar() {
          const reg = document.getElementById('regional').value;
          document.querySelector('#step2 .btn-ok').disabled = true;
          document.querySelector('#step2 .btn-cancel').disabled = true;
          document.getElementById('loading').style.display = 'block';
          
          google.script.run
            .withSuccessHandler(function() { google.script.host.close(); })
            .withFailureHandler(function(err) { 
              const loadingEl = document.getElementById('loading');
              loadingEl.textContent = "Erro: " + err.message;
              loadingEl.style.color = "red";
            })
            .processarPendencias(reg, null);
        }
      </script>
    </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Envio de Cobranças');
}

/**
 * Middleware (Filtro de Borda): Função higienizadora e unificadora de IDs de filiais.
 * Intercepta qualquer ID numérico. Se o valor for estritamente superior a 3000, 
 * o robô subtrai 3000 automaticamente em memória (num -= 3000). 
 * Isso garante que filiais 4164 ou 3001 tornem-se 1164 ou 1 antes do cruzamento de dados,
 * mitigando a orfandade de dados e agrupando todos os colaboradores na filial raiz corretamente.
 */
function normalizarFilialId(filialIdRaw) {
  const trimmed = String(filialIdRaw).trim();
  const num = Number(trimmed);
  
  if (!isNaN(num) && trimmed !== "") {
    let idFinal = num;
    if (idFinal > 3000) {
      idFinal -= 3000;
    }
    return idFinal.toString();
  }
  
  // Se não for número (ex: sigla ou vazio), retorna o próprio texto tratado
  return trimmed;
}

/**
 * Lógica principal para cruzar dados e enviar e-mails.
 * Pode ser acionada tanto pelo modal UI (Regional) quanto pelo gatilho automático (Diretoria).
 * @param {string} regionalEscolhida A regional selecionada.
 * @param {string} diretoriaEscolhida A diretoria que sofrerá o filtro (usado na automação).
 */
function processarPendencias(regionalEscolhida, diretoriaEscolhida) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const abaRol = ss.getSheetByName("Rol");
  const abaBase = ss.getSheetByName("Base");
  const abaDadosLoja = ss.getSheetByName("DADOS_LOJAS"); 

  if (!abaRol || !abaBase || !abaDadosLoja) {
    if(!diretoriaEscolhida) {
       mostrarAlertaHTML("Erro Crítico", "Uma das abas não foi encontrada. Verifique os nomes (Rol, Base, DADOS_LOJAS).");
    }
    return;
  }

  // Otimização O(1): Leitura em batch (lote) para não exceder cota de API
  const dadosRol = abaRol.getDataRange().getValues().slice(1);
  const dadosBase = abaBase.getDataRange().getValues().slice(1);
  const dadosLoja = abaDadosLoja.getDataRange().getValues().slice(1);

  // 1. Criar Set de concluintes O(1) na memória
  const idsConcluidos = new Set();
  dadosBase.forEach(row => {
    const idStr = String(row[1]).trim();
    if (idStr) idsConcluidos.add(idStr);
  });

  // 2. Mapear e-mails dos gerentes com a chave de filial higienizada pelo Middleware
  const mapaEmailsGerentes = {};
  dadosLoja.forEach(row => {
    const filialIdOriginal = row[0];
    if (filialIdOriginal) {
       const filialIdTratada = normalizarFilialId(filialIdOriginal); 
       const email = String(row[4]).trim();
       if (email) {
          mapaEmailsGerentes[filialIdTratada] = email;
       }
    }
  });

  // 3. Agrupar pendentes por filial e filtrar elegíveis
  const filiaisComPendencias = {};
  let totalPendentes = 0;
  
  // Blacklist de cargos configurável visando manutenção (estrutura O(1))
  const cargosIgnorados = new Set([
    "ASSISTENTE DE LOJA", 
    "MONTADOR"
  ]);

  dadosRol.forEach(row => {
    const id = String(row[1]).trim(); // Col B
    const nome = String(row[2]).trim(); // Col C
    const filialIdRaw = row[3]; // Col D
    const diretoria = String(row[4]).trim().toUpperCase(); // Col E
    const regional = String(row[5]).trim().toUpperCase(); // Col F
    const situacao = String(row[6]).trim(); // Col G
    const cargo = String(row[7]).trim().toUpperCase(); // Col H

    let atendeCriterioFiltro = false;
    
    if (regionalEscolhida && regional === regionalEscolhida) {
      atendeCriterioFiltro = true;
    }
    if (diretoriaEscolhida && diretoria === diretoriaEscolhida) {
      atendeCriterioFiltro = true;
    }

    // Regra: Somente ativos, id válido e que não pertençam aos cargos ignorados
    if (atendeCriterioFiltro && situacao === "Em Atividade Normal" && id) {
      if (!cargosIgnorados.has(cargo)) {
        if (!idsConcluidos.has(id)) {
          // Aplica o Middleware de Borda: intercepta filiais e agrupa chaves unificadas
          const filialIdNum = normalizarFilialId(filialIdRaw);
  
          if (!filiaisComPendencias[filialIdNum]) {
            filiaisComPendencias[filialIdNum] = [];
          }
          
          filiaisComPendencias[filialIdNum].push({ id: id, nome: nome });
          totalPendentes++;
        }
      }
    }
  });

  if (totalPendentes === 0) {
    if(!diretoriaEscolhida) {
       mostrarAlertaHTML("Aviso", `Nenhuma pendência encontrada para a seleção informada com os critérios vigentes.`);
    }
    return;
  }

  // 4. Disparar e-mails
  let emailsEnviados = 0;
  let errosEnvio = [];

  for (const filialId in filiaisComPendencias) {
    const pendentes = filiaisComPendencias[filialId];
    const emailGerente = mapaEmailsGerentes[filialId];

    if (!emailGerente) {
      errosEnvio.push(`Filial ${filialId}: E-mail não localizado na aba DADOS_LOJAS.`);
      continue;
    }

    let tabelaHtml = `<table border="1" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;">
                        <tr style="background-color: #f2f2f2;">
                          <th style="padding: 8px; text-align: left;">ID</th>
                          <th style="padding: 8px; text-align: left;">Nome</th>
                        </tr>`;
    
    pendentes.forEach(p => {
      tabelaHtml += `<tr>
                      <td style="padding: 8px;">${p.id}</td>
                      <td style="padding: 8px;">${p.nome}</td>
                     </tr>`;
    });
    tabelaHtml += `</table>`;

    const assunto = `Pendências Curso do Mês - Filial ${filialId}`;
    const corpoHtml = `
      <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.6;">
        <p>Olá Gerente da Filial <strong>${filialId}</strong>,</p>
        <p>Abaixo está a lista dos colaboradores ativos da sua loja que <strong>ainda não concluíram</strong> o curso obrigatório deste mês.</p>
        <p>Pedimos a gentileza de orientá-los a acessar o link abaixo para finalizar o treinamento o mais breve possível:</p>
        <p style="text-align: center; margin: 20px 0;">
          <a href="https://universidadeluiza.com.br" style="background-color: #0288d1; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">Acessar Universidade Luiza</a>
        </p>
        ${tabelaHtml}
        <br>
        <p>Atenciosamente,<br>Equipe de Gestão de Pessoas - Lojas</p>
      </div>
    `;

    try {
      MailApp.sendEmail({
        to: emailGerente,
        subject: assunto,
        htmlBody: corpoHtml
      });
      emailsEnviados++;
    } catch (e) {
      errosEnvio.push(`Erro (Filial ${filialId}): ${e.message}`);
    }
  }

  if (!diretoriaEscolhida) {
    let msgFinal = `<strong>Processo concluído!</strong><br><br>E-mails enviados para filiais unificadas: <strong>${emailsEnviados}</strong>`;
    if (errosEnvio.length > 0) {
      msgFinal += `<br><br><span style="color:red;"><strong>Atenção, erros em ${errosEnvio.length} filiais:</strong></span><br>` + errosEnvio.join("<br>");
    }
    mostrarAlertaHTML("Relatório de Envio", msgFinal);
  }
}
