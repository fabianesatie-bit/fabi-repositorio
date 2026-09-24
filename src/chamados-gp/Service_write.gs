/**
 * ============================================================================
 * MÓDULO: Service_Write.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Service_Write.gs)
 * ============================================================================
 */

/**
 * Roteador WebApp de entrada HTTP GET com navegação segura entre telas.
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

  if (pagina === "faq" || pagina === "ajuda") {
    return HtmlService.createTemplateFromFile('Faq')
      .evaluate()
      .setTitle('Central de Ajuda & Status dos Sistemas - GP360')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Padrão / Rota "index"
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Abertura de Chamados - GP360')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * GERENCIADOR DE FAQ VIA DASHBOARD
 */
function salvarItemFaq(itemFaq, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Acesso não autorizado." };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { sucesso: false, mensagem: "Servidor ocupado." }; }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaFaq = obterAbaFlexivel(ssChamados, CONFIG.ABA_FAQ);

    if (!abaFaq) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Aba CONFIG_FAQ não encontrada na planilha." };
    }

    var dados = abaFaq.getDataRange().getValues();
    var linhaEncontrada = -1;

    if (itemFaq.id) {
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0] || "").trim() === String(itemFaq.id).trim()) {
          linhaEncontrada = i + 1;
          break;
        }
      }
    }

    if (linhaEncontrada !== -1) {
      abaFaq.getRange(linhaEncontrada, 2).setValue(itemFaq.categoria || "Geral");
      abaFaq.getRange(linhaEncontrada, 3).setValue(itemFaq.pergunta);
      abaFaq.getRange(linhaEncontrada, 4).setValue(itemFaq.resposta);
      abaFaq.getRange(linhaEncontrada, 5).setValue(parseInt(itemFaq.ordem || 1, 10));
      abaFaq.getRange(linhaEncontrada, 6).setValue(itemFaq.status || "Ativo");
    } else {
      var novoId = "FAQ-" + ("000" + (dados.length)).slice(-3);
      abaFaq.appendRow([
        novoId,
        itemFaq.categoria || "Geral",
        itemFaq.pergunta,
        itemFaq.resposta,
        parseInt(itemFaq.ordem || dados.length, 10),
        itemFaq.status || "Ativo"
      ]);
    }

    lock.releaseLock();
    return { sucesso: true, mensagem: "Dúvida/FAQ salva com sucesso!" };
  } catch (e) {
    lock.releaseLock();
    return { sucesso: false, mensagem: "Erro ao salvar FAQ: " + e.toString() };
  }
}

/**
 * GERENCIADOR DE LINHA DO TEMPO DOS SISTEMAS VIA DASHBOARD
 */
function salvarItemTimeline(itemTimeline, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Acesso não autorizado." };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { sucesso: false, mensagem: "Servidor ocupado." }; }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaTimeline = obterAbaFlexivel(ssChamados, CONFIG.ABA_TIMELINE);

    if (!abaTimeline) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Aba CONFIG_TIMELINE não encontrada na planilha." };
    }

    var dataAtual = new Date();

    abaTimeline.appendRow([
      dataAtual,
      itemTimeline.sistemaModulo || "Senior / HCM",
      itemTimeline.tituloAviso,
      itemTimeline.mensagem,
      itemTimeline.statusSistema || "✅ Estável / Normalizado",
      itemTimeline.exibir || "Sim"
    ]);

    lock.releaseLock();
    return { sucesso: true, mensagem: "Recado de sistema salvo com sucesso!" };
  } catch (e) {
    lock.releaseLock();
    return { sucesso: false, mensagem: "Erro ao salvar aviso de sistema: " + e.toString() };
  }
}

/**
 * Salva imagem Base64 no Drive.
 */
function salvarImagemDrive(base64Data, nomeArquivo, protocolo) {
  try {
    if (!base64Data) return { sucesso: false, url: "" };
    
    var partes = base64Data.split(",");
    var contentType = partes[0].split(";")[0].replace("data:", "");
    var decoded = Utilities.base64Decode(partes[1]);
    
    var nomeFinal = (protocolo || "GP") + "_" + (nomeArquivo || "anexo.png");
    var blob = Utilities.newBlob(decoded, contentType, nomeFinal);
    
    var pasta = obterOuCriarPastaAnexosDrive();
    var arquivo = pasta.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var urlPublica = "https://lh3.googleusercontent.com/d/" + arquivo.getId();
    return { sucesso: true, url: urlPublica, fileId: arquivo.getId() };
  } catch (e) {
    Logger.log("Erro em salvarImagemDrive: " + e.toString());
    return { sucesso: false, mensagem: "Erro ao salvar imagem no Drive: " + e.toString() };
  }
}

/**
 * RESERVA E CONTROLE DE CONCORRÊNCIA EM TEMPO REAL.
 */
function assumirChamado(protocolo, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { sucesso: false, mensagem: "Servidor ocupado." }; }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);

    var dados = abaBase.getDataRange().getValues();
    var linhaEncontrada = -1;
    var statusAtual = "";
    var respAtual = "";
    var ultInteracao = null;

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0] || "").trim() === String(protocolo).trim()) {
        linhaEncontrada = i + 1;
        statusAtual = String(dados[i][9] || "Novo / Aberto").trim();
        respAtual = String(dados[i][11] || "Não atribuído").trim();
        ultInteracao = dados[i][10] ? new Date(dados[i][10]) : null;
        break;
      }
    }

    if (linhaEncontrada === -1) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Protocolo não localizado." };
    }

    var dataAtual = new Date();
    var responsavelNome = obterNomeOuEmailResponsavel(permissao.email);

    if (statusAtual.indexOf("Análise") !== -1 && respAtual !== "Não atribuído" && respAtual !== responsavelNome && ultInteracao) {
      var diffMinutos = Math.floor((dataAtual - ultInteracao) / (1000 * 60));
      if (diffMinutos < 15) {
        lock.releaseLock();
        return {
          sucesso: false,
          bloqueado: true,
          mensagem: "O coordenador (" + respAtual + ") já está atendendo este chamado no momento."
        };
      }
    }

    abaBase.getRange(linhaEncontrada, 10).setValue("Em Análise");
    abaBase.getRange(linhaEncontrada, 11).setValue(dataAtual);
    abaBase.getRange(linhaEncontrada, 12).setValue(responsavelNome);

    lock.releaseLock();
    return { sucesso: true, responsavel: responsavelNome };
  } catch (err) {
    lock.releaseLock();
    return { sucesso: false, mensagem: "Erro ao assumir chamado: " + err.toString() };
  }
}

/**
 * Registra um novo chamado.
 */
function registrarNovoChamado(dadosChamado) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (e) { return { sucesso: false, mensagem: "Sistema ocupado." }; }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);

    if (!abaBase) {
      lock.releaseLock();
      return { sucesso: false, mensagem: "Aba CHAMADOS_BASE não localizada." };
    }

    var protocolo = gerarProximoProtocolo(abaBase);
    var dataAtual = new Date();
    var filialNorm = normalizarFilialId(dadosChamado.filialId);

    var descricaoCompleta = dadosChamado.descricao || "";
    if (dadosChamado.imagemBase64) {
      var resImg = salvarImagemDrive(dadosChamado.imagemBase64, "print_abertura.png", protocolo);
      if (resImg.sucesso) {
        descricaoCompleta += "\n\n📷 Anexo de Imagem: " + resImg.url;
      }
    }

    abaBase.appendRow([
      protocolo,
      dataAtual,
      String(dadosChamado.solicitanteId).trim(),
      dadosChamado.solicitanteNome,
      filialNorm,
      dadosChamado.filialNome || ("Loja " + filialNorm),
      dadosChamado.categoria,
      dadosChamado.assunto,
      descricaoCompleta,
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
        descricaoCompleta
      ]);
    }

    try {
      dispararNotificacaoNovoChamado({
        protocolo: protocolo,
        solicitanteNome: dadosChamado.solicitanteNome,
        solicitanteId: dadosChamado.solicitanteId,
        filialId: filialNorm,
        categoria: dadosChamado.categoria,
        assunto: dadosChamado.assunto,
        descricao: descricaoCompleta
      });
    } catch (errEmail) {
      Logger.log("Aviso: E-mail não enviado: " + errEmail.toString());
    }

    lock.releaseLock();
    return { sucesso: true, protocolo: protocolo, solicitanteId: dadosChamado.solicitanteId };
  } catch (err) {
    lock.releaseLock();
    return { sucesso: false, mensagem: "Erro ao registrar chamado: " + err.toString() };
  }
}

/**
 * Registra a resposta de um Coordenador.
 */
function responderChamado(protocolo, textoResposta, marcarConcluido, chaveAcesso, imagemBase64) {
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
    var responsavelAtendimento = obterNomeOuEmailResponsavel(permissao.email);

    var textoFinal = textoResposta || "";
    if (imagemBase64) {
      var resImg = salvarImagemDrive(imagemBase64, "print_resposta.png", protocolo);
      if (resImg.sucesso) {
        textoFinal += "\n\n📷 Imagem Anexada pelo GP: " + resImg.url;
      }
    }

    abaBase.getRange(linhaEncontrada, 10).setValue(novoStatus);
    abaBase.getRange(linhaEncontrada, 11).setValue(dataAtual);
    abaBase.getRange(linhaEncontrada, 12).setValue(responsavelAtendimento);

    if (abaHist) {
      abaHist.appendRow([
        protocolo,
        dataAtual,
        responsavelAtendimento,
        "Coordenador/GP",
        textoFinal
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
 * Resolve o nome exibível do responsável.
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

  if (emailFinal && emailFinal.indexOf("@") !== -1) {
    var parteLocal = emailFinal.split("@")[0];
    var partesNome = parteLocal.split(".");
    return partesNome.map(function(p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    }).join(" ");
  }

  return emailFinal || "Fabiane Satie";
}

/**
 * Dispara 1 único e-mail com todos em cópia (CC).
 */
function dispararNotificacaoNovoChamado(info) {
  if (typeof EMAILS_AUTORIZADOS_COORD === 'undefined' || !EMAILS_AUTORIZADOS_COORD || EMAILS_AUTORIZADOS_COORD.length === 0) {
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
