/**
 * ============================================================================
 * MÓDULO: Service_Write.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Service_Write.gs)
 * ============================================================================
 */

/**
 * Roteador WebApp de entrada HTTP GET.
 */
function doGet(e) {
  var pagina = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase().trim() : "index";
  var chaveAcesso = (e && e.parameter && e.parameter.acesso) ? e.parameter.acesso.toLowerCase().trim() : "";

  if (pagina === "painel") {
    var permissao = validarAcessoCoordenador(chaveAcesso);
    if (permissao.autorizado) {
      return HtmlService.createTemplateFromFile('Painel_Coord')
        .evaluate()
        .setTitle('Dashboard & Atendimento GP360')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      return HtmlService.createHtmlOutput(
        "<div style='font-family:sans-serif; text-align:center; padding:50px; background:#FFF1F2; color:#9F1239; border-radius:12px; margin:40px auto; max-width:500px; border:1px solid #FECDD3;'>" +
        "<h2>⛔ Acesso Restrito ao GP</h2>" +
        "<p>Sua conta de e-mail (" + permissao.email + ") não possui permissão para acessar o painel de atendimento.</p>" +
        "</div>"
      ).setTitle("Acesso Restrito - GP360");
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Abertura de Chamados - GP360')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Gera o próximo número de protocolo sequencial para o ano vigente.
 * Estrutura: GP-[ANO]-[NÚMERO COM 5 DÍGITOS] (Ex: GP-2026-00001)
 *
 * @param {Sheet} abaBase Objeto da aba CHAMADOS_BASE
 * @return {string}
 */
function gerarProximoProtocolo(abaBase) {
  var anoAtual = new Date().getFullYear();
  var prefixo = "GP-" + anoAtual + "-";
  var ultimoNum = 0;

  if (abaBase) {
    try {
      var dados = abaBase.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        var prot = String(dados[i][0] || "").trim();
        if (prot.indexOf(prefixo) === 0) {
          var numPart = parseInt(prot.replace(prefixo, ""), 10);
          if (!isNaN(numPart) && numPart > ultimoNum) {
            ultimoNum = numPart;
          }
        }
      }
    } catch (e) {
      Logger.log("Erro ao calcular próximo protocolo: " + e.toString());
    }
  }

  var proximoNum = ultimoNum + 1;
  var strNum = ("00000" + proximoNum).slice(-5);
  return prefixo + strNum;
}

/**
 * Registra um novo chamado na planilha DEDICADA DE CHAMADOS (19QKV7...).
 * @param {Object} dadosChamado
 * @return {Object}
 */
function registrarNovoChamado(dadosChamado) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { sucesso: false, mensagem: "Sistema ocupado. Tente novamente em alguns instantes." };
  }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);

    if (!abaBase) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Aba CHAMADOS_BASE não localizada na planilha dedicada." };
    }

    var protocolo = gerarProximoProtocolo(abaBase);
    var dataAtual = new Date();
    var filialNorm = normalizarFilialId(dadosChamado.filialId);

    // Mapeamento idêntico às colunas da planilha dedicada:
    // A: Protocolo | B: Data_Abertura | C: Solicitante_ID | D: Solicitante_Nome
    // E: Filial_Id | F: Filial_Nome | G: Categoria | H: Assunto
    // I: Descricao_Inicial | J: Status | K: Ultima_Interacao | L: Responsavel_Atendimento
    abaBase.appendRow([
      protocolo,
      dataAtual,
      String(dadosChamado.solicitanteId).trim(),
      dadosChamado.solicitanteNome,
      filialNorm,
      dadosChamado.filialNome || ("Loja " + filialNorm),
      dadosChamado.categoria,
      dadosChamado.assunto,
      dadosChamado.descricao,
      "Novo / Aberto",
      dataAtual,
      "Não atribuído"
    ]);

    if (abaHist) {
      abaHist.appendRow([
        protocolo,
        dataAtual,
        dadosChamado.solicitanteNome,
        "Solicitante (Loja " + filialNorm + ")",
        dadosChamado.descricao
      ]);
    }

    // Disparo seguro de e-mail único com toda a equipe em cópia (CC)
    try {
      dispararNotificacaoNovoChamado({
        protocolo: protocolo,
        solicitanteNome: dadosChamado.solicitanteNome,
        solicitanteId: dadosChamado.solicitanteId,
        filialId: filialNorm,
        categoria: dadosChamado.categoria,
        assunto: dadosChamado.assunto,
        descricao: dadosChamado.descricao
      });
    } catch (errEmail) {
      Logger.log("Aviso: Chamado gravado com sucesso, mas o e-mail em cópia não pôde ser enviado: " + errEmail.toString());
    }

    lock.releaseLock();
    return { sucesso: true, protocolo: protocolo, solicitanteId: dadosChamado.solicitanteId };
  } catch (err) {
    lock.releaseLock();
    Logger.log("Erro em registrarNovoChamado: " + err.toString());
    return { sucesso: false, mensagem: "Erro ao registrar chamado: " + err.toString() };
  }
}

/**
 * Registra a resposta de um Coordenador na planilha dedicada e atribui o nome/e-mail do responsável na Coluna L.
 *
 * @param {string} protocolo
 * @param {string} textoResposta
 * @param {boolean} marcarConcluido
 * @param {string} [chaveAcesso]
 * @return {Object}
 */
function responderChamado(protocolo, textoResposta, marcarConcluido, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Permissão negada." };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { sucesso: false, mensagem: "Servidor ocupado." }; }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);

    var dados = abaBase.getDataRange().getValues();
    var linhaEncontrada = -1;

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0] || "").trim() === String(protocolo).trim()) {
        linhaEncontrada = i + 1;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Protocolo não localizado." };
    }

    var dataAtual = new Date();
    var novoStatus = marcarConcluido ? "Respondido / Concluído" : "Em Análise";

    // Captura o nome ou e-mail da pessoa logada
    var responsavelAtendimento = obterNomeOuEmailResponsavel(permissao.email);

    // Atualiza Coluna J (Status = Coluna 10), Coluna K (Ultima_Interacao = Coluna 11), Coluna L (Responsavel = Coluna 12)
    abaBase.getRange(linhaEncontrada, 10).setValue(novoStatus);
    abaBase.getRange(linhaEncontrada, 11).setValue(dataAtual);
    abaBase.getRange(linhaEncontrada, 12).setValue(responsavelAtendimento);

    if (abaHist) {
      abaHist.appendRow([
        protocolo,
        dataAtual,
        responsavelAtendimento,
        "Coordenador/GP",
        textoResposta
      ]);
    }

    lock.releaseLock();
    return { sucesso: true, protocolo: protocolo, novoStatus: novoStatus, responsavel: responsavelAtendimento };
  } catch (err) {
    lock.releaseLock();
    return { sucesso: false, mensagem: "Erro ao salvar resposta: " + err.toString() };
  }
}

/**
 * Função auxiliar para resolver o nome amigável ou o e-mail limpo de quem está respondendo.
 * @param {string} emailPermissao
 * @return {string}
 */
function obterNomeOuEmailResponsavel(emailPermissao) {
  var emailSessao = "";
  try {
    emailSessao = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (e) {
    emailSessao = "";
  }

  var emailFinal = emailSessao || String(emailPermissao || "").toLowerCase().trim();

  if (emailFinal.indexOf("fabiane.satie") !== -1) return "Fabiane Satie";
  if (emailFinal.indexOf("lorena.castro") !== -1) return "Lorena Castro";
  if (emailFinal.indexOf("milena.ramos") !== -1) return "Milena Ramos";
  if (emailFinal.indexOf("marcia.almeida") !== -1) return "Márcia Almeida";
  if (emailFinal.indexOf("josiane.emilio") !== -1) return "Josiane Emílio";

  // Se houver e-mail válido, formata o primeiro e último nome
  if (emailFinal && emailFinal.indexOf("@") !== -1) {
    var parteLocal = emailFinal.split("@")[0];
    var partesNome = parteLocal.split(".");
    var nomeFormatado = partesNome.map(function(p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    }).join(" ");
    return nomeFormatado;
  }

  return emailFinal || "Fabiane Satie";
}

/**
 * Dispara 1 único e-mail de alerta com TODOS os gestores/coordenadores em CÓPIA (CC).
 * @param {Object} info
 */
function dispararNotificacaoNovoChamado(info) {
  if (typeof EMAILS_AUTORIZADOS_COORD === 'undefined' || !EMAILS_AUTORIZADOS_COORD || EMAILS_AUTORIZADOS_COORD.length === 0) {
    Logger.log("Aviso: Lista EMAILS_AUTORIZADOS_COORD não localizada para envio de notificação.");
    return;
  }

  var listaEmails = [];
  for (var i = 0; i < EMAILS_AUTORIZADOS_COORD.length; i++) {
    var em = String(EMAILS_AUTORIZADOS_COORD[i] || "").trim().toLowerCase();
    if (em && listaEmails.indexOf(em) === -1) {
      listaEmails.push(em);
    }
  }

  if (listaEmails.length === 0) return;

  var destinatarioPrincipal = listaEmails[0];
  var emailsEmCopia = listaEmails.slice(1).join(",");

  var urlPainel = "https://script.google.com/macros/s/AKfycbwscX_uUQ97OOcqPKBOa-Pc_7BD3e9VyvZ3irTIfEpS0DfWvlG9WmFw6A-AnOsxMyMu/exec?page=painel&acesso=gp";
  var assuntoEmail = "🚨 Novo Chamado GP360: " + info.protocolo + " - Loja " + info.filialId;
  
  var corpoHtml = `
    <div style="font-family: Arial, sans-serif; color: #1E293B; max-width: 600px; border: 1px solid #E2E8F0; border-radius: 12px; padding: 24px; background: #FFF;">
      <h2 style="color: #7C3AED; margin-top: 0;">🚨 Novo Chamado Recebido (MG/CO • SUL)</h2>
      <p><strong>Protocolo:</strong> ${info.protocolo}</p>
      <p><strong>Solicitante:</strong> ${info.solicitanteNome} (Matrícula: ${info.solicitanteId})</p>
      <p><strong>Filial:</strong> Loja ${info.filialId}</p>
      <p><strong>Categoria:</strong> ${info.categoria} - ${info.assunto}</p>
      <div style="background: #F8FAFC; padding: 16px; border-radius: 8px; border-left: 4px solid #7C3AED; margin: 16px 0; white-space: pre-wrap;">
        ${info.descricao}
      </div>
      <div style="margin-top: 24px;">
        <a href="${urlPainel}" style="background: #7C3AED; color: #FFF; padding: 12px 20px; border-radius: 8px; text-decoration: none; font-weight: bold; display: inline-block;">
          👉 Abrir Painel de Atendimento
        </a>
      </div>
    </div>
  `;

  var opcoesEnvio = {
    to: destinatarioPrincipal,
    subject: assuntoEmail,
    htmlBody: corpoHtml
  };

  if (emailsEmCopia) {
    opcoesEnvio.cc = emailsEmCopia;
  }

  MailApp.sendEmail(opcoesEnvio);
}
