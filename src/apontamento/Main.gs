/**
 * ROTEADOR DE PÁGINAS E SERVIÇO DE DISPARO DE E-MAILS
 */

function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode ? String(e.parameter.mode) : '';
  var reqFilialId = e && e.parameter && e.parameter.filialId ? String(e.parameter.filialId) : '';
  var reqChapa = e && e.parameter && e.parameter.chapa ? String(e.parameter.chapa) : '';

  if (mode === 'certificado') {
    var templateCert = HtmlService.createTemplateFromFile('FormCertificado');
    templateCert.filialId = reqFilialId || '';
    templateCert.chapa = reqChapa || '';

    return templateCert.evaluate()
      .setTitle('Inclusão de Certificados | Magalu GP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
  }

  var templateIndex = HtmlService.createTemplateFromFile('Index');
  templateIndex.filialId = reqFilialId || '';
  templateIndex.chapa = reqChapa || '';

  return templateIndex.evaluate()
    .setTitle('Apontamento Operações de Loja | Magalu GP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover');
}

function sendEmailFilial(filialId, customNotes) {
  try {
    var rawData = obterDadosAuditoria(false);
    var data = JSON.parse(rawData);
    if (!data.sucesso) throw new Error(data.erro);

    var filialNorm = normalizarFilialId(filialId);
    var loja = data.filiaisAlertas.find(function(l) { return l.filial === filialNorm; });
    if (!loja) throw new Error('Filial não localizada.');

    var irregularidadesLoja = data.apontamentosBrutos.filter(function(a) { return a.filialId === filialNorm; });
    if (irregularidadesLoja.length === 0) {
      return { success: false, message: 'Nenhuma irregularidade encontrada para esta filial.' };
    }

    var recipient = loja.gerenteEmail || EMAIL_ORGANIZADOR_PADRAO;
    var subject = '[Magalu GP] Apontamento de Ponto Eletrônico - ' + loja.nomeLoja;
    var htmlBody = buildEmailHtml(loja, irregularidadesLoja, customNotes);

    MailApp.sendEmail({
      to: recipient,
      cc: EMAIL_ORGANIZADOR_PADRAO,
      subject: subject,
      htmlBody: htmlBody,
      name: 'Gestão de Pessoas Magalu'
    });

    return {
      success: true,
      message: 'Notificação enviada com sucesso para ' + recipient + ' (Cópia para ' + EMAIL_ORGANIZADOR_PADRAO + ').'
    };
  } catch (err) {
    Logger.log('Erro ao enviar e-mail filial: ' + err.toString());
    return { success: false, error: err.message };
  }
}

function sendEmailRegional(regionalName) {
  try {
    var rawData = obterDadosAuditoria(false);
    var data = JSON.parse(rawData);
    if (!data.sucesso) throw new Error(data.erro);

    var lojasRegional = data.filiaisAlertas.filter(function(l) {
      return regionalName === 'TODAS' || l.regional === regionalName;
    });

    var disparados = 0;
    lojasRegional.forEach(function(loja) {
      var irregularidades = data.apontamentosBrutos.filter(function(a) { return a.filialId === loja.filial; });
      if (irregularidades.length > 0) {
        var recipient = loja.gerenteEmail || EMAIL_ORGANIZADOR_PADRAO;
        var subject = '[Magalu GP] Apontamento de Ponto Eletrônico - ' + loja.nomeLoja;
        var htmlBody = buildEmailHtml(loja, irregularidades, 'Disparo automático em massa da Regional ' + loja.regional);

        MailApp.sendEmail({
          to: recipient,
          cc: EMAIL_ORGANIZADOR_PADRAO,
          subject: subject,
          htmlBody: htmlBody,
          name: 'Gestão de Pessoas Magalu'
        });
        disparados++;
      }
    });

    return {
      success: true,
      message: 'Disparo em massa concluído! E-mails enviados para ' + disparados + ' filial(is) em alerta.'
    };
  } catch (err) {
    Logger.log('Erro ao enviar disparo regional: ' + err.toString());
    return { success: false, error: err.message };
  }
}

function buildEmailHtml(loja, irregularidades, customNotes) {
  var mapaAgrupado = {};

  irregularidades.forEach(function(item) {
    var key = (item.chapa || item.nome) + '_' + item.tipoIrregularidade;
    if (!mapaAgrupado[key]) {
      mapaAgrupado[key] = {
        chapa: item.chapa || '-',
        nome: item.nome,
        cargo: item.cargo,
        tipoIrregularidade: item.tipoIrregularidade,
        qtdTotal: 0
      };
    }
    mapaAgrupado[key].qtdTotal += (item.quantidadeMes || 1);
  });

  var rows = '';
  Object.values(mapaAgrupado).forEach(function(g) {
    rows += '<tr>' +
      '<td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #1e293b;"><strong>' + g.nome + '</strong><br><span style="font-size: 11px; color: #64748b;">ID (Chapa): <strong>' + g.chapa + '</strong> | ' + g.cargo + '</span></td>' +
      '<td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-size: 12px; color: #dc2626; font-weight: bold;">' + g.tipoIrregularidade + '</td>' +
      '<td style="padding: 10px; border-bottom: 1px solid #e2e8f0; font-size: 13px; text-align: center; font-weight: bold; color: #0f172a;">' + g.qtdTotal + '</td>' +
    '</tr>';
  });

  var webAppUrl = ScriptApp.getService().getUrl();
  var linkAnexoGerente = webAppUrl + '?mode=certificado&filialId=' + loja.filial;
  var textoWhatsApp = "Olá! Identificamos apontamentos em seu registro de ponto neste mês. Solicitamos que realize a trilha obrigatória no link: " + LINK_CURSO_OBRIGATORIO + " e envie o comprovante para a gerência.";

  return '<div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">' +
    '<div style="background: linear-gradient(135deg, #0086ff, #a855f7); color: #ffffff; padding: 24px;">' +
      '<h1 style="margin: 0; font-size: 20px; font-weight: bold;">Magazine Luiza - Gestão de Pessoas</h1>' +
      '<p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Auditoria de Ponto Eletrônico e Jornada de Trabalho - Julho/2026</p>' +
    '</div>' +
    '<div style="padding: 24px;">' +
      '<p style="font-size: 15px; color: #0f172a; margin-top: 0;"><strong>Olá, Gerente!</strong></p>' +
      '<p style="font-size: 13px; color: #334155; line-height: 1.6;">Identificamos apontamentos no registro de jornada da sua filial (<strong>' + loja.nomeLoja + '</strong>). No <em>Jeito Luiza de Ser</em>, cuidamos das pessoas garantindo que o descanso e os limites da jornada sejam rigorosamente respeitados.</p>' +
      (customNotes ? '<div style="background-color: #f8fafc; border-left: 4px solid #0086ff; padding: 12px; margin: 16px 0; font-size: 13px; color: #1e293b;"><strong>Observações do Coordenador de GP:</strong><br>' + customNotes + '</div>' : '') +
      '<h3 style="font-size: 14px; color: #0f172a; margin: 20px 0 10px; border-bottom: 2px solid #0086ff; padding-bottom: 4px;">Colaboradores com Inconsistências na Filial</h3>' +
      '<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">' +
        '<thead>' +
          '<tr style="background-color: #f1f5f9; text-align: left;">' +
            '<th style="padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569;">ID / Colaborador</th>' +
            '<th style="padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569;">Inconsistência</th>' +
            '<th style="padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569; text-align: center;">Qtd Total</th>' +
          '</tr>' +
        '</thead>' +
        '<tbody>' + rows + '</tbody>' +
      '</table>' +

      '<div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 16px; margin: 20px 0;">' +
        '<h4 style="margin: 0 0 6px; font-size: 13px; color: #1e40af;">🎓 1. Treinamento Obrigatório do Colaborador</h4>' +
        '<p style="margin: 0 0 12px; font-size: 12px; color: #1e3a8a; line-height: 1.5;">Encaminhe o link do treinamento da Universidade Luiza para os colaboradores afetados:</p>' +
        '<a href="' + LINK_CURSO_OBRIGATORIO + '" style="display: inline-block; background-color: #0086ff; color: #ffffff; padding: 10px 16px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold; margin-bottom: 12px;" target="_blank">Acessar Curso: Ponto Eletrônico - Guia Essencial</a><br>' +

        '<div style="background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 12px; margin-top: 10px;">' +
          '<p style="margin: 0 0 6px; font-size: 11px; font-weight: bold; color: #1e40af;">Caro gerente, encaminhe esse texto para o colaborador realizar o curso na universidade luiza:</p>' +
          '<code style="font-size: 11px; color: #0f172a; display: block; background: #f8fafc; padding: 8px; border-radius: 4px; border: 1px solid #e2e8f0; line-height: 1.4;">' + textoWhatsApp + '</code>' +
        '</div>' +

        '<div style="border-top: 1px dashed #bfdbfe; padding-top: 14px; margin-top: 14px;">' +
          '<h4 style="margin: 0 0 6px; font-size: 13px; color: #065f46;">📎 2. Envio do Certificado pelo Gerente de Loja</h4>' +
          '<p style="margin: 0 0 10px; font-size: 12px; color: #1e3a8a;">Após a conclusão do treinamento pelos colaboradores, clique no botão abaixo para abrir a página de inclusão dos comprovantes:</p>' +
          '<a href="' + linkAnexoGerente + '" style="display: inline-block; background-color: #10b981; color: #ffffff; padding: 10px 18px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold;" target="_blank">Inclusão de Certificados pelo Gerente</a>' +
        '</div>' +
      '</div>' +

      '<p style="font-size: 12px; color: #64748b; margin-bottom: 0;">Dúvidas ou alinhamentos? Responda a este e-mail ou acione a equipe de GP Lojas em <a href="mailto:' + EMAIL_ORGANIZADOR_PADRAO + '" style="color: #0086ff;">' + EMAIL_ORGANIZADOR_PADRAO + '</a>.</p>' +
    '</div>' +
    '<div style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px 24px; text-align: center; font-size: 11px; color: #94a3b8;">' +
      'Magazine Luiza S.A. • Gestão de Pessoas • SLA de Acompanhamento: 15 Dias' +
    '</div>' +
  '</div>';
}
