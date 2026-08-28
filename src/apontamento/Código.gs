/**
 * @file Code.gs
 * @description Google Apps Script Backend para o Dashboard de Auditoria de Ponto Eletrônico Magalu
 * @author Gestão de Pessoas (GP) - Magazine Luiza S.A.
 * @version 1.0.0
 */

// IDs das Planilhas Oficiais
const SPREADSHEET_AUDITORIA_ID = '1RHbEZ67n9ZjieKiMcpm_qqTVmFsmjyv9nIfDqpwm0WM';
const SPREADSHEET_ACESSOS_LOJAS_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA';

// Cargos com permissão de acesso ao dashboard
const CARGOS_AUTORIZADOS = [
  'GERENTERH',
  'E-mail Geral Área',
  'DiretorRH',
  'Diretor OP',
  'Coordenador',
  'Regional OP'
];

const LINK_CURSO_OBRIGATORIO = 'https://universidadeluiza.com.br/app/home/canal/magalu?section=csc-conte-sempre-comigo&trail=ponto-eletronico-o-guia-essencial-do-colaborador';
const EMAIL_ORGANIZADOR_PADRAO = 'gplojas@magazineluiza.com.br';

/**
 * Ponto de entrada do Web App
 */
function doGet(e) {
  try {
    const userEmail = Session.getActiveUser().getEmail() || 'gplojas@magazineluiza.com.br';
    const accessInfo = checkUserAccess(userEmail);
    
    const template = HtmlService.createTemplateFromFile('Index');
    template.userEmail = userEmail;
    template.isAuthorized = accessInfo.authorized;
    template.userProfile = JSON.stringify(accessInfo.profile || {});
    
    return template.evaluate()
      .setTitle('Auditoria de Ponto Magalu')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('<h3>Erro ao carregar o sistema: ' + error.message + '</h3>');
  }
}

/**
 * Validação de acesso por E-mail e Cargo
 */
function checkUserAccess(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ACESSOS_LOJAS_ID);
    const sheetUsers = ss.getSheetByName('DADOS_USUARIOS');
    if (!sheetUsers) throw new Error('Aba DADOS_USUARIOS não encontrada.');

    const data = sheetUsers.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    
    const idxEmail = headers.indexOf('Email');
    const idxNome = headers.indexOf('Nome');
    const idxCargo = headers.indexOf('Cargo');
    const idxDiretoria = headers.indexOf('Diretoria_Principal');
    const idxRegionais = headers.indexOf('Regionais_Atendidas');
    const idxNivel = headers.indexOf('Nivel_Acesso');

    const cleanEmail = String(email).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][idxEmail]).trim().toLowerCase();
      if (rowEmail === cleanEmail) {
        const cargo = String(data[i][idxCargo]).trim();
        const authorized = CARGOS_AUTORIZADOS.some(c => c.toLowerCase() === cargo.toLowerCase());
        
        return {
          authorized: authorized,
          profile: {
            email: rowEmail,
            nome: data[i][idxNome],
            cargo: cargo,
            diretoriaPrincipal: data[i][idxDiretoria],
            regionaisAtendidas: String(data[i][idxRegionais]).split(',').map(r => r.trim()),
            nivelAcesso: data[i][idxNivel]
          }
        };
      }
    }

    // Caso não encontrado na base
    return { authorized: false, profile: null, message: 'Usuário não cadastrado na base de acessos.' };
  } catch (err) {
    Logger.log('Erro em checkUserAccess: ' + err.toString());
    return { authorized: false, error: err.toString() };
  }
}

/**
 * Consolidação dos dados das 3 abas analíticas + Lojas
 */
function getDashboardData() {
  try {
    const userEmail = Session.getActiveUser().getEmail() || 'gplojas@magazineluiza.com.br';
    const access = checkUserAccess(userEmail);
    if (!access.authorized) {
      throw new Error('Acesso negado: Perfil não possui permissão para visualizar estes dados.');
    }

    // 1. Carregar Dados de Lojas
    const ssLojas = SpreadsheetApp.openById(SPREADSHEET_ACESSOS_LOJAS_ID);
    const sheetLojas = ssLojas.getSheetByName('DADOS_LOJAS');
    const lojasData = sheetLojas.getDataRange().getValues();
    const lojasHeaders = lojasData[0].map(h => String(h).trim());
    
    const idxFilialId = lojasHeaders.indexOf('Filial_ID');
    const idxNomeFantasia = lojasHeaders.indexOf('Nome_Fantasia');
    const idxLojasRegional = lojasHeaders.indexOf('Regional');
    const idxLojasDiretoria = lojasHeaders.indexOf('Diretoria');
    const idxGerenteEmail = lojasHeaders.indexOf('Gerente_Email');
    const idxCidade = lojasHeaders.indexOf('Cidade');
    const idxEstado = lojasHeaders.indexOf('Estado');

    const lojasMap = {};
    const listaLojas = [];

    for (let i = 1; i < lojasData.length; i++) {
      const fid = String(lojasData[i][idxFilialId]).trim();
      const objLoja = {
        filialId: fid,
        nomeFantasia: lojasData[i][idxNomeFantasia] || ('Filial ' + fid),
        regional: lojasData[i][idxLojasRegional],
        diretoria: lojasData[i][idxLojasDiretoria],
        gerenteEmail: lojasData[i][idxGerenteEmail],
        cidade: lojasData[i][idxCidade],
        estado: lojasData[i][idxEstado]
      };
      lojasMap[fid] = objLoja;
      listaLojas.push(objLoja);
    }

    // 2. Carregar Abas da Planilha Analítica
    const ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    const apontamentos = [];

    // Aba 1: Acesso fora da jornada
    const sheetAcesso = ssAuditoria.getSheetByName('Acesso fora da jornada JULHO/2026');
    if (sheetAcesso) {
      const dataAcesso = sheetAcesso.getDataRange().getValues();
      const headers = dataAcesso[0].map(h => String(h).trim());
      const idxF = headers.indexOf('FILIAL');
      const idxNome = headers.indexOf('NOME');
      const idxCargo = headers.indexOf('CARGO');
      const idxData = headers.indexOf('DATA_BATIDA') !== -1 ? headers.indexOf('DATA_BATIDA') : headers.indexOf('DATA_TRANSACAO');
      const idxVendaRemota = headers.indexOf('VENDA REMOTA?');
      const idxReg = headers.indexOf('REGIONAL');
      const idxDir = headers.indexOf('DIRETOR');

      for (let i = 1; i < dataAcesso.length; i++) {
        const fid = String(dataAcesso[i][idxF]).trim();
        const lojaInfo = lojasMap[fid] || {
          nomeFantasia: 'Filial ' + fid,
          regional: dataAcesso[i][idxReg] || 'Regional Não Cadastrada',
          diretoria: dataAcesso[i][idxDir] || 'Diretoria Geral'
        };

        apontamentos.push({
          id: 'AFJ-' + i,
          filialId: fid,
          filialNome: lojaInfo.nomeFantasia,
          regional: lojaInfo.regional,
          diretoria: lojaInfo.diretoria,
          chapa: String(dataAcesso[i][headers.indexOf('CDICONTRATADO')] || 'N/D'),
          nome: dataAcesso[i][idxNome] || 'Colaborador',
          cargo: dataAcesso[i][idxCargo] || 'Vendedor',
          tipoIrregularidade: 'Acesso Fora da Jornada',
          subtipo: dataAcesso[i][idxVendaRemota] === 'Sim' ? 'Venda Remota Fora de Jornada' : 'Acesso ao PDV/Sistema',
          dataOcorrencia: dataAcesso[i][idxData] ? String(dataAcesso[i][idxData]).substring(0, 10) : '2026-07-15',
          quantidadeMes: Number(dataAcesso[i][headers.indexOf('QTDE BATIDAS')]) || 1
        });
      }
    }

    // Aba 2: Horas extras
    const sheetHE = ssAuditoria.getSheetByName('Horas extras JULHO/2026');
    if (sheetHE) {
      const dataHE = sheetHE.getDataRange().getValues();
      const headersHE = dataHE[0].map(h => String(h).trim());
      const idxF = headersHE.indexOf('FILIAL_OFICIAL');
      const idxNome = headersHE.indexOf('NOME');
      const idxCargo = headersHE.indexOf('CARGO');
      const idxData = headersHE.indexOf('DATA_OCORRENCIA');
      const idxHoras = headersHE.indexOf('QTDE_HORAS');

      for (let i = 1; i < dataHE.length; i++) {
        const fid = String(dataHE[i][idxF]).trim();
        const lojaInfo = lojasMap[fid] || {
          nomeFantasia: 'Filial ' + fid,
          regional: 'Regional Sul',
          diretoria: 'Diretoria Lojas Sul/Sudeste'
        };

        apontamentos.push({
          id: 'HE-' + i,
          filialId: fid,
          filialNome: lojaInfo.nomeFantasia,
          regional: lojaInfo.regional,
          diretoria: lojaInfo.diretoria,
          chapa: String(dataHE[i][headersHE.indexOf('CHAPA')] || 'N/D'),
          nome: dataHE[i][idxNome] || 'Colaborador',
          cargo: dataHE[i][idxCargo] || 'Colaborador',
          tipoIrregularidade: 'Horas Extras Não Autorizadas',
          subtipo: 'Excesso de Jornada (' + (dataHE[i][idxHoras] || '0') + 'h)',
          dataOcorrencia: dataHE[i][idxData] ? String(dataHE[i][idxData]).substring(0, 10) : '2026-07-20',
          quantidadeMes: 1,
          detalhesAdicionais: { qtdeHoras: String(dataHE[i][idxHoras] || '0') }
        });
      }
    }

    // Aba 3: Ajuste / Britânicos
    const sheetBritanico = ssAuditoria.getSheetByName('Ajuste / Britânicos JULHO/2026');
    if (sheetBritanico) {
      const dataB = sheetBritanico.getDataRange().getValues();
      const headersB = dataB[0].map(h => String(h).trim());
      const idxF = headersB.indexOf('UNIDADE');
      const idxNome = headersB.indexOf('NOME');
      const idxCargo = headersB.indexOf('CARGO');
      const idxAjustes = headersB.indexOf('QTDE_AJUSTES');
      const idxPerc = headersB.indexOf('PERC_BRITANICO');

      for (let i = 1; i < dataB.length; i++) {
        const fid = String(dataB[i][idxF]).trim();
        const lojaInfo = lojasMap[fid] || {
          nomeFantasia: 'Filial ' + fid,
          regional: dataB[i][headersB.indexOf('NOME_REGIONAL')] || 'Regional Magalu',
          diretoria: dataB[i][headersB.indexOf('DIRETOR')] || 'Diretoria Geral'
        };

        apontamentos.push({
          id: 'BRIT-' + i,
          filialId: fid,
          filialNome: lojaInfo.nomeFantasia,
          regional: lojaInfo.regional,
          diretoria: lojaInfo.diretoria,
          chapa: String(dataB[i][headersB.indexOf('CDI')] || 'N/D'),
          nome: dataB[i][idxNome] || 'Colaborador',
          cargo: dataB[i][idxCargo] || 'Colaborador',
          tipoIrregularidade: 'Ajuste / Marcação Britânica',
          subtipo: 'Marcação Invariável (' + (dataB[i][idxPerc] || '0%') + ')',
          dataOcorrencia: '2026-07-25',
          quantidadeMes: Number(dataB[i][idxAjustes]) || 1,
          detalhesAdicionais: { percBritanico: String(dataB[i][idxPerc] || '0%') }
        });
      }
    }

    return {
      success: true,
      currentUser: access.profile,
      lojas: listaLojas,
      apontamentos: apontamentos
    };
  } catch (error) {
    Logger.log('Erro em getDashboardData: ' + error.toString());
    return { success: false, error: error.message };
  }
}

/**
 * Envio de E-mail Individual para o Gerente da Filial
 */
function sendEmailFilial(filialId, customNotes) {
  try {
    const data = getDashboardData();
    if (!data.success) throw new Error(data.error);

    const loja = data.lojas.find(l => l.filialId === String(filialId));
    if (!loja) throw new Error('Filial não localizada.');

    const irregularidadesLoja = data.apontamentos.filter(a => a.filialId === String(filialId));
    if (irregularidadesLoja.length === 0) {
      return { success: false, message: 'Nenhuma irregularidade encontrada para esta filial.' };
    }

    const recipient = loja.gerenteEmail;
    if (!recipient || recipient.indexOf('@') === -1) {
      throw new Error('E-mail do gerente da filial não cadastrado ou inválido em DADOS_LOJAS.');
    }

    const subject = '[Magalu GP] Auditoria de Ponto e Jornada - ' + loja.nomeFantasia;
    const htmlBody = buildEmailHtml(loja, irregularidadesLoja, customNotes);

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

/**
 * Envio de E-mail em Massa para todos os gerentes da Regional
 */
function sendEmailRegional(regionalName) {
  try {
    const data = getDashboardData();
    if (!data.success) throw new Error(data.error);

    const lojasRegional = data.lojas.filter(l => l.regional === regionalName);
    let enviados = 0;
    const erros = [];

    lojasRegional.forEach(loja => {
      const items = data.apontamentos.filter(a => a.filialId === loja.filialId);
      if (items.length > 0 && loja.gerenteEmail) {
        try {
          const subject = '[Magalu GP - Regional ' + regionalName + '] Auditoria de Ponto - ' + loja.nomeFantasia;
          const htmlBody = buildEmailHtml(loja, items, 'Disparo automático de fechamento mensal regional de jornada.');

          MailApp.sendEmail({
            to: loja.gerenteEmail,
            cc: EMAIL_ORGANIZADOR_PADRAO,
            subject: subject,
            htmlBody: htmlBody,
            name: 'Gestão de Pessoas Magalu'
          });
          enviados++;
        } catch (e) {
          erros.push(loja.nomeFantasia + ': ' + e.message);
        }
      }
    });

    return {
      success: true,
      enviados: enviados,
      erros: erros,
      message: 'Processamento concluído: ' + enviados + ' e-mails disparados com sucesso para a Regional ' + regionalName + '.'
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Construtor do Template HTML de E-mail Oficial Magalu
 */
function buildEmailHtml(loja, irregularidades, customNotes) {
  let tableRows = '';
  irregularidades.forEach((item, index) => {
    const bg = index % 2 === 0 ? '#f8f9ff' : '#ffffff';
    tableRows += '<tr style="background-color: ' + bg + ';">' +
      '<td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; font-weight: 600; color: #1e293b;">' + item.nome + '</td>' +
      '<td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #475569;">' + item.cargo + '</td>' +
      '<td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #dc2626; font-weight: 500;">' + item.tipoIrregularidade + '</td>' +
      '<td style="padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; text-align: center; font-weight: 700; color: #0086ff;">' + item.quantidadeMes + '</td>' +
    '</tr>';
  });

  const notesSection = customNotes ? 
    '<div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 12px 16px; margin: 16px 0; border-radius: 4px; font-size: 13px; color: #92400e;"><strong>Nota da Coordenação GP:</strong> ' + customNotes + '</div>' : '';

  return '<!DOCTYPE html>' +
  '<html>' +
  '<head><meta charset="utf-8"></head>' +
  '<body style="font-family: Arial, sans-serif; background-color: #f1f5f9; padding: 24px 0; margin: 0;">' +
    '<div style="max-width: 650px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">' +
      // Header Magalu
      '<div style="background: linear-gradient(90deg, #0086ff 0%, #a855f7 100%); padding: 24px; text-align: left; color: #ffffff;">' +
        '<h2 style="margin: 0; font-size: 20px; font-weight: 700;">Auditoria de Ponto e Jornada de Trabalho</h2>' +
        '<p style="margin: 4px 0 0 0; font-size: 13px; opacity: 0.9;">Gestão de Pessoas (GP) | Magazine Luiza S.A.</p>' +
      '</div>' +
      // Body
      '<div style="padding: 24px;">' +
        '<p style="font-size: 15px; color: #1e293b; margin-top: 0;"><strong>Olá, Gerente!</strong></p>' +
        '<p style="font-size: 14px; color: #475569; line-height: 1.6;">' +
          'Esperamos que este e-mail o(a) encontre bem. No Magalu, nosso maior patrimônio são as nossas pessoas e o respeito rigoroso aos seus direitos e à jornada de trabalho.' +
        '</p>' +
        '<p style="font-size: 14px; color: #475569; line-height: 1.6;">' +
          'Apresentamos abaixo o consolidado mensal de colaboradores da <strong>' + loja.nomeFantasia + '</strong> (' + loja.cidade + '/' + loja.estado + ') que apresentaram inconsistências ou irregularidades em seus registros de ponto no ciclo recente:' +
        '</p>' +
        notesSection +
        // Tabela
        '<table style="width: 100%; border-collapse: collapse; margin: 20px 0; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">' +
          '<thead>' +
            '<tr style="background-color: #eff4ff; color: #1e40af; text-align: left;">' +
              '<th style="padding: 10px 12px; font-size: 12px; text-transform: uppercase;">Colaborador</th>' +
              '<th style="padding: 10px 12px; font-size: 12px; text-transform: uppercase;">Cargo</th>' +
              '<th style="padding: 10px 12px; font-size: 12px; text-transform: uppercase;">Tipo de Irregularidade</th>' +
              '<th style="padding: 10px 12px; font-size: 12px; text-transform: uppercase; text-align: center;">Qtd no Mês</th>' +
            '</tr>' +
          '</thead>' +
          '<tbody>' +
            tableRows +
          '</tbody>' +
        '</table>' +
        // Instrução Obrigatória
        '<div style="background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 16px; margin: 24px 0;">' +
          '<p style="margin: 0; font-size: 13px; color: #1e3a8a; line-height: 1.6; font-weight: 500;">' +
            '📌 <strong>Instrução Obrigatória:</strong><br>' +
            'Solicitamos que os colaboradores notificados concluam o treinamento obrigatório de ponto eletrônico no link abaixo e enviem os certificados anexados em resposta a este e-mail:<br>' +
            '<a href="' + LINK_CURSO_OBRIGATORIO + '" target="_blank" style="color: #0086ff; font-weight: 700; text-decoration: underline; display: inline-block; margin-top: 6px;">' +
              '👉 Acessar Treinamento: Ponto Eletrônico - O Guia Essencial do Colaborador (Universidade Luiza)' +
            '</a>' +
          '</p>' +
        '</div>' +
        '<p style="font-size: 13px; color: #64748b; line-height: 1.5;">' +
          'Contamos com a sua liderança calorosa e próxima para orientar a equipe no Jeito Luiza de Ser. Caso tenha dúvidas, responda diretamente a este e-mail da Coordenação de GP.' +
        '</p>' +
        '<p style="font-size: 13px; color: #334155; margin-bottom: 0;">' +
          'Um caloroso abraço,<br><strong>Coordenação de Gestão de Pessoas (GP Lojas)</strong><br>' +
          '<span style="color: #94a3b8; font-size: 11px;">Magazine Luiza S.A. | ' + EMAIL_ORGANIZADOR_PADRAO + '</span>' +
        '</p>' +
      '</div>' +
    '</div>' +
  '</body>' +
  '</html>';
}
