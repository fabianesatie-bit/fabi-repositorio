// =============================================================================
// CONFIGURAÇÕES GERAIS DE IDS CORPORATIVOS
// =============================================================================
const SPREADSHEET_ID = '1tn2FiNVWVMFM-3DC14_L0LaHGmd9O-teU2cvsIoajkk';
const TEMPLATE_DOC_ID = '1kOrTd6yUrTQ3TXJiAO0ZOwMMM9g94VS_-mIdcvQXfuM';
const TEMPLATE_DESLIGAMENTO_DOC_ID = '1nulmYsjwTu-diseZzTUlg8gT5s81Jk10Mtjc1ts7cqY'; 
const DRIVE_FOLDER_ID = '150Br7ERz-bNtyesrUUQ2sxRCFnJPEE1y';

// =============================================================================
// HELPERS E CACHE DE MEMÓRIA
// =============================================================================
const SpreadsheetCache = {};

/**
 * Normaliza o ID da filial com base na regra de negócio corporativa (> 3000).
 */
function normalizarFilialId(id) {
  if (id === null || id === undefined) return '';
  let idStr = id.toString().trim();
  if (idStr === '') return '';
  
  if (idStr.includes('.')) {
    idStr = idStr.split('.')[0];
  }
  
  let num = parseInt(idStr.replace(/\D/g, ''), 10);
  if (isNaN(num)) return idStr;
  
  if (num > 3000) {
    num -= 3000;
  }
  return num.toString();
}

function getCachedSheetData(spreadsheetId, sheetName) {
  const cacheKey = spreadsheetId + '_' + sheetName;
  if (SpreadsheetCache[cacheKey]) return SpreadsheetCache[cacheKey];
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    if (!ss) return null;
    
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        if (normalizarTexto(sheets[i].getName()) === normalizarTexto(sheetName)) {
          sheet = sheets[i];
          break;
        }
      }
    }
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    SpreadsheetCache[cacheKey] = data;
    return data;
  } catch (err) {
    Logger.log("Erro ao abrir planilha/aba (" + sheetName + "): " + err.message);
    return null;
  }
}

function normalizarTexto(texto) {
  if (!texto) return '';
  return texto.toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") 
    .trim();
}

function obterIndiceCabecalho(headers, termosBusca, indicePadrao) {
  if (!headers || !Array.isArray(headers)) return indicePadrao;
  for (let i = 0; i < headers.length; i++) {
    const h = normalizarTexto(headers[i]);
    for (let j = 0; j < termosBusca.length; j++) {
      if (h.indexOf(normalizarTexto(termosBusca[j])) !== -1) {
        return i;
      }
    }
  }
  return indicePadrao;
}

function obterCalendarioGP() {
  let calendar = null;
  try {
    calendar = CalendarApp.getCalendarById('gplojas@magazineluiza.com.br');
  } catch (e) {}
  if (!calendar) {
    calendar = CalendarApp.getDefaultCalendar();
  }
  return calendar;
}

function aplicarPermissoesArquivo(fileId, emailsString, nivelAcesso) {
  if (!emailsString) return;
  try {
    const file = DriveApp.getFileById(fileId);
    const emails = emailsString
      .replace(/;/g, ',')
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(e => e && e.includes('@'));

    emails.forEach(email => {
      try {
        if (nivelAcesso === 'EDITOR') {
          file.addEditor(email);
        } else if (nivelAcesso === 'COMENTADOR') {
          file.addCommenter(email);
        } else {
          file.addViewer(email);
        }
      } catch (errUser) {
        Logger.log(`Falha ao conceder acesso ao usuário: ${email}. Detalhe: ${errUser.message}`);
      }
    });
  } catch (e) {
    Logger.log(`Erro crítico no motor de permissões para o arquivo ${fileId}: ${e.message}`);
  }
}

// =============================================================================
// 1. CARREGAMENTO INICIAL E ROTEAMENTO (doGet)
// =============================================================================
function doGet(e) {
  let emailLogado = '';
  try {
    emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (err) {}

  let page = e && e.parameter && e.parameter.page ? e.parameter.page : 'inicio';
  let action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  let id = e && e.parameter && e.parameter.id ? e.parameter.id : '';
  let idFeedback = e && e.parameter && e.parameter.idFeedback ? e.parameter.idFeedback : '';
  let idDesligamento = e && e.parameter && e.parameter.idDesligamento ? e.parameter.idDesligamento : '';

  const template = HtmlService.createTemplateFromFile('Index');
  template.emailLogado = emailLogado;
  template.page = page;
  template.action = action;
  template.editId = id;
  template.idFeedback = idFeedback;
  template.idDesligamento = idDesligamento;

  return template.evaluate()
    .setTitle('GP Magalu - Case-Flow & Feedback')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obterSessaoInicial(pageReq) {
  let email = '';
  try {
    email = Session.getActiveUser().getEmail().toLowerCase().trim();
  } catch (e) {}

  let autorizado = false;
  let usuario = null;

  try {
    autorizado = verificarAutorizacao(email, pageReq);
    usuario = obterUsuarioLogado(email);
  } catch (err) {
    Logger.log("Erro na validação de sessão: " + err.message);
  }

  return { email: email, autorizado: autorizado, usuario: usuario };
}

// =============================================================================
// 2. VERIFICAÇÃO DE PERMISSÃO DE ACESSO E HIERARQUIA
// =============================================================================
function verificarAutorizacao(email, pageReq) {
  if (!email) return false;
  email = email.toLowerCase().trim();

  if (pageReq === 'gerente') {
    return true;
  }

  if (email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br') {
    return true;
  }

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (!dataUsuarios) return false;

    const headers = dataUsuarios[0];
    const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
    const colHierarquia = obterIndiceCabecalho(headers, ['cargo'], 2);
    const colNivel = obterIndiceCabecalho(headers, ['nivel_acesso'], 5);

    for (let i = 1; i < dataUsuarios.length; i++) {
      const emailCadastrado = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
      if (emailCadastrado === email) {
        const hierarquia = dataUsuarios[i][colHierarquia] ? normalizarTexto(dataUsuarios[i][colHierarquia]) : '';
        const nivel = dataUsuarios[i][colNivel] ? normalizarTexto(dataUsuarios[i][colNivel]) : '';

        if (hierarquia.includes('diretor op') || hierarquia.includes('diretoria op') || hierarquia.includes('gerente regional op') || hierarquia.includes('regional op')) {
          return true;
        }

        if (nivel === 'bloqueado' || nivel === 'inativo') return false;

        if (hierarquia.includes('gp') || hierarquia.includes('coord') || hierarquia.includes('diretor')) {
          return true;
        }

        if (/gerente.*loja/i.test(hierarquia) || hierarquia === 'gerente') return false;
        
        return true; 
      }
    }
  } catch (err) {}
  return false;
}

function verificarEhAdminMaster(email) {
  if (!email) return false;
  email = email.toLowerCase().trim();

  if (email === 'fabiane.satie@magazineluiza.com.br' || email === 'gplojas@magazineluiza.com.br') {
    return true;
  }

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (!dataUsuarios) return false;

    const headers = dataUsuarios[0];
    const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
    const colCargo = obterIndiceCabecalho(headers, ['cargo'], 2);
    const colNivel = obterIndiceCabecalho(headers, ['nivel_acesso'], 5);

    for (let i = 1; i < dataUsuarios.length; i++) {
      const emailCadastrado = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
      if (emailCadastrado === email) {
        const cargo = dataUsuarios[i][colCargo] ? normalizarTexto(dataUsuarios[i][colCargo]) : '';
        const nivel = dataUsuarios[i][colNivel] ? normalizarTexto(dataUsuarios[i][colNivel]) : '';
        
        if (nivel.includes('admin') || nivel.includes('master') || cargo.includes('coord') || cargo.includes('gerentegp') || cargo.includes('gerente gp') || cargo.includes('diretor')) {
          return true;
        }
      }
    }
  } catch (e) {}
  return false;
}

function obterUsuarioLogado(email) {
  if (!email) return { email: '', nome: 'COORDENADOR GP', apelido: 'GP', cargo: 'GP', ehRegionalOuDiretor: false };
  let apelido = email.split('@')[0];
  if (apelido.includes('.')) {
    const parts = apelido.split('.');
    apelido = parts[0].charAt(0).toUpperCase() + parts[0].slice(1);
  } else {
    apelido = apelido.charAt(0).toUpperCase() + apelido.slice(1);
  }

  let cargoCadastrado = 'GP';
  let ehRegionalOuDiretor = false;

  try {
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    if (dataUsuarios && dataUsuarios.length > 1) {
      const headers = dataUsuarios[0];
      const colEmail = obterIndiceCabecalho(headers, ['email'], 0);
      const colCargo = obterIndiceCabecalho(headers, ['cargo'], 2);

      for (let i = 1; i < dataUsuarios.length; i++) {
        const emailCad = dataUsuarios[i][colEmail] ? dataUsuarios[i][colEmail].toString().toLowerCase().trim() : '';
        if (emailCad === email.toLowerCase().trim()) {
          cargoCadastrado = dataUsuarios[i][colCargo] ? dataUsuarios[i][colCargo].toString().trim() : 'GP';
          const cargoNorm = normalizarTexto(cargoCadastrado);
          if (cargoNorm.includes('gerente regional op') || cargoNorm.includes('regional op') || cargoNorm.includes('diretor op') || cargoNorm.includes('diretoria op')) {
            ehRegionalOuDiretor = true;
          }
          break;
        }
      }
    }
  } catch (e) {}

  return {
    email: email,
    nome: email.split('@')[0].replace('.', ' ').toUpperCase(),
    apelido: apelido,
    cargo: cargoCadastrado,
    ehRegionalOuDiretor: ehRegionalOuDiretor
  };
}

// =============================================================================
// 3. BUSCA RELACIONAL DE CONTATOS E GERENTE DA FILIAL (DADOS_ATIVOS)
// =============================================================================
function obterContatosPorFilial(filial) {
  try {
    const dataLojas = getCachedSheetData(SPREADSHEET_ID, 'DADOS_LOJAS');
    const dataUsuarios = getCachedSheetData(SPREADSHEET_ID, 'DADOS_USUARIOS');
    const dataAtivos = getCachedSheetData(SPREADSHEET_ID, 'DADOS_ATIVOS');
    
    if (!dataLojas || !dataUsuarios) return null;

    const filialNorm = normalizarFilialId(filial);

    const headersLojas = dataLojas[0];
    const colFilial = obterIndiceCabecalho(headersLojas, ['filial_id', 'filial'], 0);
    const colRegLoja = obterIndiceCabecalho(headersLojas, ['regional'], 2);
    const colDirLoja = obterIndiceCabecalho(headersLojas, ['diretoria'], 3);
    const colGerEmail = obterIndiceCabecalho(headersLojas, ['gerente_email', 'gerente'], 4);

    let regionalFilial = '', diretoriaFilial = '', gerenteLojaEmail = '', encontrouLoja = false;

    for (let i = 1; i < dataLojas.length; i++) {
      const fVal = dataLojas[i][colFilial] ? dataLojas[i][colFilial].toString().trim() : '';
      if (normalizarFilialId(fVal) === filialNorm) {
        regionalFilial = dataLojas[i][colRegLoja] ? dataLojas[i][colRegLoja].toString().trim() : '';
        diretoriaFilial = dataLojas[i][colDirLoja] ? dataLojas[i][colDirLoja].toString().trim() : '';
        gerenteLojaEmail = dataLojas[i][colGerEmail] ? dataLojas[i][colGerEmail].toString().trim() : '';
        encontrouLoja = true;
        break;
      }
    }

    if (!encontrouLoja) return null;

    let gerenteNomeCompleto = '';
    let gerenteIdAtivo = '';
    let gerenteDisplayFormatted = '';

    if (dataAtivos && dataAtivos.length > 1) {
      const hAtivos = dataAtivos[0];
      const colAtivFilial = obterIndiceCabecalho(hAtivos, ['filial'], 3);
      const colAtivId = obterIndiceCabecalho(hAtivos, ['id'], 1);
      const colAtivNome = obterIndiceCabecalho(hAtivos, ['nome'], 2);
      const colAtivSit = obterIndiceCabecalho(hAtivos, ['situacao'], 6);
      const colAtivCargoRes = obterIndiceCabecalho(hAtivos, ['cargo_resumido', 'cargo resumido'], 9);

      for (let k = 1; k < dataAtivos.length; k++) {
        const row = dataAtivos[k];
        const fAtiv = row[colAtivFilial] ? normalizarFilialId(row[colAtivFilial]) : '';
        
        if (fAtiv === filialNorm) {
          const cargoRes = row[colAtivCargoRes] ? normalizarTexto(row[colAtivCargoRes]) : '';
          const sit = row[colAtivSit] ? normalizarTexto(row[colAtivSit]) : '';

          const ehGerente = cargoRes.includes('gerente');
          const statusValido = sit.includes('atividade normal') || sit.includes('ferias') || sit.includes('gozando ferias');

          if (ehGerente && statusValido) {
            gerenteIdAtivo = row[colAtivId] ? row[colAtivId].toString().trim() : '';
            gerenteNomeCompleto = row[colAtivNome] ? row[colAtivNome].toString().trim().toUpperCase() : '';
            
            if (gerenteIdAtivo && gerenteNomeCompleto) {
              gerenteDisplayFormatted = gerenteIdAtivo + " - " + gerenteNomeCompleto;
            } else {
              gerenteDisplayFormatted = gerenteNomeCompleto;
            }
            break;
          }
        }
      }
    }

    if (!gerenteDisplayFormatted) {
      gerenteDisplayFormatted = gerenteLojaEmail || "Gerente não localizado na base ativa";
    }

    const headersUsers = dataUsuarios[0];
    const uColEmail = obterIndiceCabecalho(headersUsers, ['email'], 0);
    const uColHierarquia = obterIndiceCabecalho(headersUsers, ['cargo'], 2);
    const uColRegional = obterIndiceCabecalho(headersUsers, ['regionais_atendidas', 'regionais'], 4); 
    const uColDiretoria = obterIndiceCabecalho(headersUsers, ['diretoria_principal', 'diretoria'], 3); 
    
    let coordenador = [], gerenteGP = [], regionalOp = [], compliance = [], diretorRH = [], diretorOp = [];
    const normRegFilial = normalizarTexto(regionalFilial);
    const normDirFilial = normalizarTexto(diretoriaFilial);

    for (let i = 1; i < dataUsuarios.length; i++) {
      const uEmail = dataUsuarios[i][uColEmail] ? dataUsuarios[i][uColEmail].toString().trim() : '';
      if (!uEmail || uEmail.indexOf('@') === -1) continue;

      const uHierarquia = dataUsuarios[i][uColHierarquia] ? normalizarTexto(dataUsuarios[i][uColHierarquia]) : '';
      const uReg = dataUsuarios[i][uColRegional] ? normalizarTexto(dataUsuarios[i][uColRegional]) : '';
      const uDir = dataUsuarios[i][uColDiretoria] ? normalizarTexto(dataUsuarios[i][uColDiretoria]) : '';

      const ehCoord = uHierarquia.includes('coord');
      const ehGerenteGP = uHierarquia.includes('gerentegp') || uHierarquia.includes('gerente gp');
      const ehRegionalOp = uHierarquia.includes('gerente regional op') || uHierarquia.includes('regional op') || (uHierarquia.includes('regional') && !uHierarquia.includes('gp'));
      const ehDiretorOp = uHierarquia.includes('diretor op') || uHierarquia.includes('diretoria op');
      const ehDiretorRH = uHierarquia.includes('diretorrh') || uHierarquia.includes('diretor rh');
      const ehCompliance = uHierarquia.includes('compli') || uHierarquia.includes('etica');

      const matchReg = (uReg === 'todas' || uReg === '' || normRegFilial.includes(uReg) || uReg.includes(normRegFilial));
      const matchDir = (uDir === 'todas' || uDir === '' || normDirFilial.includes(uDir) || uDir.includes(normDirFilial));

      if (ehCoord && matchReg) coordenador.push(uEmail);
      if (ehRegionalOp && matchReg) regionalOp.push(uEmail);
      if (ehGerenteGP && matchDir) gerenteGP.push(uEmail);
      if (ehDiretorOp && (matchDir || matchReg)) diretorOp.push(uEmail);
      if (ehCompliance) compliance.push(uEmail);
      if (ehDiretorRH) diretorRH.push(uEmail);
    }

    return {
      filial: filialNorm,
      filialOriginal: filial,
      diretoria: diretoriaFilial,
      regional: regionalFilial,
      coordenador: coordenador.join(','),
      gerenteGP: gerenteGP.join(','),
      regionalEmail: regionalOp.join(','),
      diretorRH: diretorRH.join(','),
      compliance: compliance.join(','),
      gerenteLoja: gerenteDisplayFormatted,
      gerenteEmail: gerenteLojaEmail,
      gerenteIdAtivo: gerenteIdAtivo,
      gerenteNomeAtivo: gerenteNomeCompleto,
      diretorOp: diretorOp.join(',')
    };
  } catch (err) {
    Logger.log("Erro contatos: " + err);
  }
  return null;
}

function buscarColaboradorAtivo(idColaborador) {
  try {
    const data = getCachedSheetData(SPREADSHEET_ID, 'DADOS_ATIVOS');
    if (!data) return null;

    const headers = data[0];
    const colId = obterIndiceCabecalho(headers, ['id'], 1);
    const colNome = obterIndiceCabecalho(headers, ['nome'], 2);
    const colCargo = obterIndiceCabecalho(headers, ['cargo'], 7);
    const colTempo = obterIndiceCabecalho(headers, ['tempo_empresa'], 4);
    const colFilial = obterIndiceCabecalho(headers, ['filial'], 3);

    for (let i = 1; i < data.length; i++) {
      if (data[i][colId] && data[i][colId].toString().trim() === idColaborador.toString().trim()) {
        return {
          nome: data[i][colNome] ? data[i][colNome].toString() : '',
          cargo: data[i][colCargo] ? data[i][colCargo].toString() : '',
          tempo: data[i][colTempo] ? data[i][colTempo].toString() : '',
          filial: data[i][colFilial] ? normalizarFilialId(data[i][colFilial]) : ''
        };
      }
    }
  } catch (err) {}
  return null;
}

function salvarEvidenciasDrive(arquivosBase64, filial) {
  if (!arquivosBase64 || arquivosBase64.length === 0) return '';
  const links = [];
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    arquivosBase64.forEach((arq, index) => {
      if (arq.dados && arq.dados.includes(',')) {
        const split = arq.dados.split(',');
        const contentType = split[0].match(/:(.*?);/)[1];
        const rawData = Utilities.base64Decode(split[1]);
        const fileName = "Evidencia_F" + filial + "_" + new Date().getTime() + "_" + (arq.nome || ("Print_" + index + ".png"));
        const file = folder.createFile(Utilities.newBlob(rawData, contentType, fileName));
        links.push(file.getUrl());
      } else if (arq.dados && arq.dados.startsWith('http')) {
        links.push(arq.dados);
      }
    });
  } catch(e) {
    Logger.log("Erro salvando evidencias: " + e.message);
  }
  return links.join(' \n');
}

// =============================================================================
// 4. CONSULTA E GESTÃO INTERNA DE REGISTROS (READ, LIST, SOFT DELETE)
// =============================================================================
function listarTodosRegistrosUsuario() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const lista = [];

    const dataAp = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataAp && dataAp.length > 1) {
      for (let i = 1; i < dataAp.length; i++) {
        const emailCriador = dataAp[i][32] ? dataAp[i][32].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          let dt = dataAp[i][1];
          if (dt instanceof Date) dt = dt.toLocaleDateString('pt-BR');
          lista.push({
            tipo: 'apuracao',
            tipoLabel: 'Apuração',
            id: dataAp[i][0] ? dataAp[i][0].toString() : '',
            data: dt ? dt.toString() : 'N/A',
            filial: dataAp[i][3] ? dataAp[i][3].toString() : '',
            status: dataAp[i][7] ? dataAp[i][7].toString() : 'Concluído',
            resumo: dataAp[i][13] ? dataAp[i][13].toString() : 'N/A',
            emailCriador: emailCriador
          });
        }
      }
    }

    const dataInt = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataInt && dataInt.length > 1) {
      for (let i = 1; i < dataInt.length; i++) {
        const emailCriador = dataInt[i][16] ? dataInt[i][16].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          lista.push({
            tipo: 'intervencao',
            tipoLabel: 'Feedback Clima',
            id: dataInt[i][0] ? dataInt[i][0].toString() : '',
            data: 'N/A',
            filial: dataInt[i][1] ? dataInt[i][1].toString() : '',
            status: dataInt[i][3] ? dataInt[i][3].toString() : 'Registrado',
            resumo: dataInt[i][4] ? dataInt[i][4].toString() : 'N/A',
            emailCriador: emailCriador
          });
        }
      }
    }

    const dataDes = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_DESLIGAMENTO_F');
    if (dataDes && dataDes.length > 1) {
      for (let i = 1; i < dataDes.length; i++) {
        const emailCriador = dataDes[i][21] ? dataDes[i][21].toString().toLowerCase().trim() : '';
        if (ehMaster || emailCriador === emailLogado) {
          let dt = dataDes[i][1];
          if (dt instanceof Date) dt = dt.toLocaleDateString('pt-BR');
          lista.push({
            tipo: 'desligamento',
            tipoLabel: 'Desligamento',
            id: dataDes[i][0] ? dataDes[i][0].toString() : '',
            data: dt ? dt.toString() : 'N/A',
            filial: dataDes[i][2] ? dataDes[i][2].toString() : '',
            status: dataDes[i][12] ? dataDes[i][12].toString() : 'Pendente',
            resumo: "Colaborador: " + (dataDes[i][3] ? dataDes[i][3].toString() : 'N/A'),
            emailCriador: emailCriador
          });
        }
      }
    }

    return lista;
  } catch (e) {
    return [];
  }
}

function buscarRegistroParaEdicao(id) {
  try {
    if (!id) return { erro: 'ID do registro não fornecido para edição.' };
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const idBuscado = id.toString().trim().toUpperCase();

    const dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      for (let i = 1; i < dataApuracao.length; i++) {
        if (dataApuracao[i][0] && dataApuracao[i][0].toString().trim().toUpperCase() === idBuscado) {
          const emailCriador = dataApuracao[i][32] ? dataApuracao[i][32].toString().toLowerCase().trim() : '';
          
          if (!ehMaster && emailCriador && emailCriador !== emailLogado) {
            return { erro: 'Acesso Negado: Apenas o criador original deste registro pode editá-lo.' };
          }

          let dataRec = dataApuracao[i][11];
          if (dataRec instanceof Date) dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          
          let dataFin = dataApuracao[i][12];
          if (dataFin instanceof Date) dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');

          return {
            tipo: 'apuracao',
            id: dataApuracao[i][0].toString(),
            dataRegistro: dataApuracao[i][1] ? dataApuracao[i][1].toString() : '',
            origem: dataApuracao[i][2] ? dataApuracao[i][2].toString() : 'Canal',
            filial: dataApuracao[i][3] ? dataApuracao[i][3].toString() : '',
            diretoria: dataApuracao[i][4] ? dataApuracao[i][4].toString() : '',
            regional: dataApuracao[i][5] ? dataApuracao[i][5].toString() : '',
            apurador: dataApuracao[i][6] ? dataApuracao[i][6].toString() : '',
            conclusao: dataApuracao[i][7] ? dataApuracao[i][7].toString() : '',
            linkDoc: dataApuracao[i][8] ? dataApuracao[i][8].toString() : '',
            notaHumorAntes: dataApuracao[i][9] ? dataApuracao[i][9].toString() : '',
            tratativa: dataApuracao[i][10] ? dataApuracao[i][10].toString() : '',
            dataRecebimento: dataRec ? dataRec.toString() : '',
            dataFinalizacao: dataFin ? dataFin.toString() : '',
            resumo: dataApuracao[i][13] ? dataApuracao[i][13].toString() : '',
            diagnostico: dataApuracao[i][14] ? dataApuracao[i][14].toString() : '',
            justificativa: dataApuracao[i][15] ? dataApuracao[i][15].toString() : '',
            enviarFeedbackGerente: dataApuracao[i][16] ? dataApuracao[i][16].toString() : 'nao',
            feedbackGerenteText: dataApuracao[i][17] ? dataApuracao[i][17].toString() : '',
            gerenteId: dataApuracao[i][18] ? dataApuracao[i][18].toString() : '',
            gerenteNome: dataApuracao[i][19] ? dataApuracao[i][19].toString() : '',
            agendarIntervencao: dataApuracao[i][20] ? dataApuracao[i][20].toString() : 'nao',
            detalhesIntervencao: dataApuracao[i][21] ? dataApuracao[i][21].toString() : '',
            denunciados: [
              { id: dataApuracao[i][22] ? dataApuracao[i][22].toString() : '', nome: dataApuracao[i][23] ? dataApuracao[i][23].toString() : '', filial: dataApuracao[i][24] ? dataApuracao[i][24].toString() : '' },
              { id: dataApuracao[i][25] ? dataApuracao[i][25].toString() : '', nome: dataApuracao[i][26] ? dataApuracao[i][26].toString() : '', filial: dataApuracao[i][27] ? dataApuracao[i][27].toString() : '' },
              { id: dataApuracao[i][28] ? dataApuracao[i][28].toString() : '', nome: dataApuracao[i][29] ? dataApuracao[i][29].toString() : '', filial: dataApuracao[i][30] ? dataApuracao[i][30].toString() : '' }
            ],
            emailCriador: emailCriador
          };
        }
      }
    }

    const dataDesligamento = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_DESLIGAMENTO_F');
    if (dataDesligamento && dataDesligamento.length > 1) {
      for (let i = 1; i < dataDesligamento.length; i++) {
        if (dataDesligamento[i][0] && dataDesligamento[i][0].toString().trim().toUpperCase() === idBuscado) {
          const emailCriador = dataDesligamento[i][21] ? dataDesligamento[i][21].toString().toLowerCase().trim() : '';

          if (!ehMaster && emailCriador && emailCriador !== emailLogado) {
            return { erro: 'Acesso Negado: Apenas o criador original deste registro pode editá-lo.' };
          }

          return {
            tipo: 'desligamento',
            id: dataDesligamento[i][0].toString(),
            filial: dataDesligamento[i][2] ? dataDesligamento[i][2].toString() : '',
            colaboradorNome: dataDesligamento[i][3] ? dataDesligamento[i][3].toString() : '',
            colaboradorId: dataDesligamento[i][4] ? dataDesligamento[i][4].toString() : '',
            tempoEmpresa: dataDesligamento[i][5] ? dataDesligamento[i][5].toString() : '',
            colaboradorCargo: dataDesligamento[i][6] ? dataDesligamento[i][6].toString() : '',
            resultados: dataDesligamento[i][7] ? dataDesligamento[i][7].toString() : '',
            justificativa: dataDesligamento[i][8] ? dataDesligamento[i][8].toString() : '',
            evidencias: dataDesligamento[i][9] ? dataDesligamento[i][9].toString() : '',
            parecerCoordenador: dataDesligamento[i][10] ? dataDesligamento[i][10].toString() : '',
            coordenadorNome: dataDesligamento[i][18] ? dataDesligamento[i][18].toString() : '',
            linkDoc: dataDesligamento[i][20] ? dataDesligamento[i][20].toString() : '',
            emailCriador: emailCriador
          };
        }
      }
    }

    const dataIntervencao = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataIntervencao && dataIntervencao.length > 1) {
      for (let i = 1; i < dataIntervencao.length; i++) {
        if (dataIntervencao[i][0] && dataIntervencao[i][0].toString().trim().toUpperCase() === idBuscado) {
          const emailCriador = dataIntervencao[i][16] ? dataIntervencao[i][16].toString().toLowerCase().trim() : '';

          if (!ehMaster && emailCriador && emailCriador !== emailLogado) {
            return { erro: 'Acesso Negado: Apenas o criador original deste registro pode editá-lo.' };
          }

          let dataProg = dataIntervencao[i][6];
          if (dataProg instanceof Date) dataProg = Utilities.formatDate(dataProg, Session.getScriptTimeZone(), 'yyyy-MM-dd');

          return {
            tipo: 'intervencao',
            id: dataIntervencao[i][0].toString(),
            filial: dataIntervencao[i][1] ? dataIntervencao[i][1].toString() : '',
            linkDoc: dataIntervencao[i][2] ? dataIntervencao[i][2].toString() : '',
            statusEvolucao: dataIntervencao[i][3] ? dataIntervencao[i][3].toString() : '',
            detalhes: dataIntervencao[i][4] ? dataIntervencao[i][4].toString() : '',
            notaHumorDepois: dataIntervencao[i][5] ? dataIntervencao[i][5].toString() : '',
            novaDataProgramada: dataProg ? dataProg.toString() : '',
            colaboradores: [
              { id: dataIntervencao[i][7] ? dataIntervencao[i][7].toString() : '', nome: dataIntervencao[i][8] ? dataIntervencao[i][8].toString() : '', filial: dataIntervencao[i][9] ? dataIntervencao[i][9].toString() : '' },
              { id: dataIntervencao[i][10] ? dataIntervencao[i][10].toString() : '', nome: dataIntervencao[i][11] ? dataIntervencao[i][11].toString() : '', filial: dataIntervencao[i][12] ? dataIntervencao[i][12].toString() : '' },
              { id: dataIntervencao[i][13] ? dataIntervencao[i][13].toString() : '', nome: dataIntervencao[i][14] ? dataIntervencao[i][14].toString() : '', filial: dataIntervencao[i][15] ? dataIntervencao[i][15].toString() : '' }
            ],
            emailCriador: emailCriador
          };
        }
      }
    }

    return { erro: 'Registro ' + id + ' não localizado nas bases.' };
  } catch (e) {
    return { erro: 'Erro ao buscar registro: ' + e.toString() };
  }
}

function cancelarRegistroProcesso(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ehMaster = verificarEhAdminMaster(emailLogado);
    const idBuscado = payload.id.toString().trim().toUpperCase();
    const motivo = payload.motivo || 'Lançado Incorretamente / Anulado';

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    let sheet = ss.getSheetByName('HISTORICO_APURACAO');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][32] ? vals[i][32].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este registro.' };
          }
          sheet.getRange(i + 1, 8).setValue('CANCELADO');
          sheet.getRange(i + 1, 16).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    sheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][21] ? vals[i][21].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este registro.' };
          }
          sheet.getRange(i + 1, 13).setValue('CANCELADO');
          sheet.getRange(i + 1, 16).setValue('CANCELADO');
          sheet.getRange(i + 1, 9).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    sheet = ss.getSheetByName('Intervencoes_Feedback');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][16] ? vals[i][16].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este registro.' };
          }
          sheet.getRange(i + 1, 4).setValue('CANCELADO');
          sheet.getRange(i + 1, 5).setValue('CANCELADO / MOTIVO: ' + motivo);
          return { sucesso: true };
        }
      }
    }

    sheet = ss.getSheetByName('Rascunhos_Apuracoes');
    if (sheet) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idBuscado) {
          const criador = vals[i][31] ? vals[i][31].toString().toLowerCase().trim() : '';
          if (!ehMaster && criador && criador !== emailLogado) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador pode anular este rascunho.' };
          }

          let apurSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
          apurSheet.appendRow([
            vals[i][0], vals[i][1], vals[i][2], vals[i][3], vals[i][4], vals[i][5], vals[i][6], 'CANCELADO',
            vals[i][8], vals[i][9], vals[i][10], vals[i][11], vals[i][12], vals[i][13],
            vals[i][14], 'CANCELADO / MOTIVO: ' + motivo, vals[i][16], vals[i][17],
            vals[i][18], vals[i][19], vals[i][20], vals[i][21],
            vals[i][22], vals[i][23], vals[i][24],
            vals[i][25], vals[i][26], vals[i][27],
            vals[i][28], vals[i][29], vals[i][30],
            new Date().toLocaleString('pt-BR'), criador
          ]);

          sheet.deleteRow(i + 1);
          return { sucesso: true };
        }
      }
    }

    return { sucesso: false, erro: 'Registro não localizado para anulação.' };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// 5. MÁQUINA DE ESTADOS E RASCUNHOS DE APURAÇÃO
// =============================================================================
function buscarRascunhosApuracoes() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const usuarioLogadoObj = obterUsuarioLogado(emailLogado);
    const ehMaster = verificarEhAdminMaster(emailLogado);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Rascunhos_Apuracoes');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const rascunhos = [];
    for (let i = 1; i < data.length; i++) {
      const emailCriador = data[i][31] ? data[i][31].toString().toLowerCase().trim() : '';
      const apuradorNome = data[i][6] ? data[i][6].toString().toUpperCase().trim() : '';

      if (!ehMaster) {
        const bateuEmail = emailCriador && (emailCriador === emailLogado);
        const bateuNome = (!emailCriador) && apuradorNome && (apuradorNome === usuarioLogadoObj.nome.toUpperCase());
        if (!bateuEmail && !bateuNome) {
          continue;
        }
      }

      let dataReg = data[i][1];
      if (dataReg instanceof Date) dataReg = dataReg.toLocaleDateString('pt-BR');
      
      let dataRec = data[i][11];
      if (dataRec instanceof Date) dataRec = Utilities.formatDate(dataRec, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      let dataFin = data[i][12];
      if (dataFin instanceof Date) dataFin = Utilities.formatDate(dataFin, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      rascunhos.push({
        idRascunho: data[i][0] ? data[i][0].toString() : '',
        dataRegistro: dataReg ? dataReg.toString() : '',
        origem: data[i][2] ? data[i][2].toString() : 'Canal',
        filial: data[i][3] ? data[i][3].toString() : '',
        apurador: data[i][6] ? data[i][6].toString() : 'GP',
        conclusao: data[i][7] ? data[i][7].toString() : 'Rascunho Pessoal',
        linkDoc: data[i][8] ? data[i][8].toString() : '',
        notaHumorAntes: data[i][9] ? data[i][9].toString() : '',
        dataRecebimento: dataRec ? dataRec.toString() : '',
        dataFinalizacao: dataFin ? dataFin.toString() : '',
        resumo: data[i][13] ? data[i][13].toString() : '',
        diagnostico: data[i][14] ? data[i][14].toString() : '',
        justificativa: data[i][15] ? data[i][15].toString() : '',
        tratativa: data[i][10] ? data[i][10].toString() : '',
        enviarFeedbackGerente: data[i][16] ? data[i][16].toString() : 'nao',
        feedbackGerenteText: data[i][17] ? data[i][17].toString() : '',
        gerenteId: data[i][18] ? data[i][18].toString() : '',
        gerenteNome: data[i][19] ? data[i][19].toString() : '',
        agendarIntervencao: data[i][20] ? data[i][20].toString() : 'nao',
        detalhesIntervencao: data[i][21] ? data[i][21].toString() : '',
        denunciados: [
          { id: data[i][22] ? data[i][22].toString() : '', nome: data[i][23] ? data[i][23].toString() : '', filial: data[i][24] ? data[i][24].toString() : '' },
          { id: data[i][25] ? data[i][25].toString() : '', nome: data[i][26] ? data[i][26].toString() : '', filial: data[i][27] ? data[i][27].toString() : '' },
          { id: data[i][28] ? data[i][28].toString() : '', nome: data[i][29] ? data[i][29].toString() : '', filial: data[i][30] ? data[i][30].toString() : '' }
        ],
        emailCriador: emailCriador
      });
    }
    return rascunhos;
  } catch (e) {
    return [];
  }
}

function salvarRascunhoApuracao(dados, tipoAcao) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Rascunhos_Apuracoes') || ss.insertSheet('Rascunhos_Apuracoes');
    
    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);
    const diretoria = contatos ? contatos.diretoria : '';
    const regional = contatos ? contatos.regional : '';
    const dataAtual = new Date().toLocaleDateString('pt-BR');
    
    let idRascunho = dados.idRascunho;
    let rowIndex = -1;
    let linkDocExistente = '';
    
    if (idRascunho) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0].toString() === idRascunho.toString()) {
          rowIndex = i + 1;
          linkDocExistente = vals[i][8] ? vals[i][8].toString() : '';
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      idRascunho = 'RASC-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const nomesDenunciadosArray = dados.denunciados.map(d => d.nome).filter(n => n);
    const stringDenunciadosUnificada = nomesDenunciadosArray.join(', ') || 'Não Informado';

    let linkDoc = linkDocExistente;
    let statusEstado = (tipoAcao === 'comite') ? 'Sob Análise Comitê' : 'Rascunho Pessoal';

    if (tipoAcao === 'comite') {
      linkDoc = criarOuAtualizarDocPreliminarComite(idRascunho, linkDocExistente, dados, filial, diretoria, regional, stringDenunciadosUnificada, contatos);
    }
    
    const rowData = [
      idRascunho, dataAtual, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', statusEstado,
      linkDoc, dados.nota_humor_antes || '', dados.tratativa || '', dados.data_recebimento || '', dados.data_finalizacao || '', dados.resumo || '',
      dados.diagnostico || '', dados.justificativa || '', dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente || '',
      dados.gerente_id || '', dados.gerente_nome || '', dados.agendar_intervencao || 'nao', dados.detalhes_intervencao || '',
      dados.denunciados[0]?.id || '', dados.denunciados[0]?.nome || '', dados.denunciados[0]?.filial || '',
      dados.denunciados[1]?.id || '', dados.denunciados[1]?.nome || '', dados.denunciados[1]?.filial || '',
      dados.denunciados[2]?.id || '', dados.denunciados[2]?.nome || '', dados.denunciados[2]?.filial || '',
      emailLogado
    ];
    
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    if (tipoAcao === 'comite' && contatos) {
      dispararAlertaComitePreliminar(contatos, filial, statusEstado, linkDoc, dados.origem, stringDenunciadosUnificada);
    }
    
    return { sucesso: true, idRascunho: idRascunho, linkDoc: linkDoc, tipoAcao: tipoAcao };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function criarOuAtualizarDocPreliminarComite(idRascunho, linkExistente, dados, filial, diretoria, regional, denunciadosStr, contatos) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    let doc = null;
    let docId = '';

    if (linkExistente && linkExistente.includes('document/d/')) {
      const match = linkExistente.match(/[-\w]{25,}/);
      if (match) docId = match[0];
      try { if (docId) doc = DocumentApp.openById(docId); } catch(e){}
    }

    if (!doc) {
      const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
      const newFileName = "[PRELIMINAR - COMITÊ] Relatorio F." + filial + " - " + denunciadosStr + " - " + new Date().toLocaleDateString('pt-BR');
      const copiedFile = templateFile.makeCopy(newFileName, folder);
      docId = copiedFile.getId();
      doc = DocumentApp.openById(docId);
    }

    const body = doc.getBody();
    body.setText(''); 

    body.appendParagraph("REGISTRO PRELIMINAR - REVISÃO DE COMITÊ").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Atenção: Este documento é um RASCUNHO EM ANÁLISE e não possui validade de encerramento definitivo.\n");

    body.appendParagraph("1. DADOS DE CONTROLE").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Filial: " + filial + " | Diretoria: " + diretoria + " | Regional: " + regional);
    body.appendParagraph("Apurador: " + (dados.apurador || 'GP'));
    body.appendParagraph("Origem: " + (dados.origem || 'Canal'));
    body.appendParagraph("Denunciado(s): " + denunciadosStr);
    body.appendParagraph("Data Recebimento: " + (dados.data_recebimento || 'N/A') + " | Previsão Conclusão: " + (dados.data_finalizacao || 'N/A'));

    body.appendParagraph("\n2. RESUMO DA DENÚNCIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(dados.resumo || 'Em preenchimento...');

    body.appendParagraph("\n3. DIAGNÓSTICO PRELIMINAR APURADO").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(dados.diagnostico || 'Em preenchimento...');

    body.appendParagraph("\n4. PARECER E JUSTIFICATIVA TEMPORÁRIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Parecer Proposto: " + (dados.conclusao || 'Sob Análise Comitê'));
    body.appendParagraph("Justificativa: " + (dados.justificativa || 'Em preenchimento...'));
    body.appendParagraph("Plano de Ação/Tratativa Proposta: " + (dados.tratativa || 'Em preenchimento...'));

    doc.saveAndClose();

    if (contatos) {
      aplicarPermissoesArquivo(docId, contatos.coordenador, 'LEITOR');
      aplicarPermissoesArquivo(docId, contatos.regionalEmail, 'LEITOR');
      if (dados.origem !== 'Interna') {
        aplicarPermissoesArquivo(docId, contatos.gerenteGP, 'LEITOR');
        aplicarPermissoesArquivo(docId, contatos.compliance, 'LEITOR');
        aplicarPermissoesArquivo(docId, contatos.diretorRH, 'LEITOR');
      }
    }

    return DriveApp.getFileById(docId).getUrl();
  } catch (e) {
    return '';
  }
}

function dispararAlertaComitePreliminar(contatos, filial, status, linkDoc, origem, denunciadosStr) {
  const lista = [];
  if (origem === 'Interna') {
    if (contatos.coordenador) contatos.coordenador.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.regionalEmail) contatos.regionalEmail.split(',').forEach(e => lista.push(e.trim()));
  } else {
    if (contatos.coordenador) contatos.coordenador.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.gerenteGP) contatos.gerenteGP.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.regionalEmail) contatos.regionalEmail.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.compliance) contatos.compliance.split(',').forEach(e => lista.push(e.trim()));
    if (contatos.diretorRH) contatos.diretorRH.split(',').forEach(e => lista.push(e.trim()));
  }

  const destinatarios = [...new Set(lista)].filter(e => e);
  if (destinatarios.length === 0) return;

  const assunto = "[ANÁLISE COMITÊ] Apuração Pendente F." + filial + " (" + denunciadosStr + ")";
  const htmlBody = 
    '<div style="font-family: Arial, sans-serif; color: #1a1a1a; max-width: 600px; margin: 0 auto; border: 1px solid #c3dafe; border-radius: 8px; overflow: hidden;">' +
      '<div style="background-color: #4C51BF; padding: 20px; color: white; text-align: center;">' +
        '<div style="background-color: white; color: #4C51BF; height: 40px; width: 40px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; margin-bottom: 10px;">📋</div>' +
        '<h2 style="margin: 0; font-size: 18px; text-transform: uppercase;">Apreciação de Comitê</h2>' +
        '<p style="margin: 4px 0 0 0; font-size: 12px; font-weight: bold; opacity: 0.9;">GP & GOVERNANÇA</p>' +
      '</div><div style="padding: 24px;">' +
        '<p style="font-size: 15px;">Um relatório de apuração da <strong>Filial ' + filial + '</strong> foi enviado para avaliação e parecer do Comitê de Ética / GP.</p>' +
        '<div style="background-color: #EBF8FF; padding: 18px; border-radius: 8px; border-left: 4px solid #3182CE; margin: 20px 0;">' +
          '<p style="margin:0 0 8px 0;"><strong>Envolvido(s):</strong> ' + denunciadosStr + '</p>' +
          '<p style="margin:0;"><strong>Origem:</strong> ' + origem + '</p>' +
        '</div>' +
        '<div style="text-align: center; margin: 30px 0;"><a href="' + linkDoc + '" target="_blank" rel="noopener noreferrer" style="background-color: #4C51BF; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold;">Visualizar Rascunho do Comitê</a></div>' +
      '</div></div>';

  MailApp.sendEmail({ to: destinatarios.join(','), subject: assunto, htmlBody: htmlBody });
}

// =============================================================================
// 6. PROCESSAMENTO DE NOVA APURAÇÃO DEFINITIVA OU EDIÇÃO (UPDATE IN-PLACE)
// =============================================================================
function processarNovaApuracao(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    
    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);

    let diretoria = 'Não identificada', regional = 'Não identificada';
    let emailCoord = '', emailGerGP = '', emailRegGP = '', emailGerenteLoja = '', emailCompliance = '', emailDirRH = '', emailDirOp = '';

    if (contatos) {
      diretoria = contatos.diretoria || diretoria;
      regional = contatos.regional || regional;
      emailCoord = contatos.coordenador;
      emailGerGP = contatos.gerenteGP;
      emailRegGP = contatos.regionalEmail;
      emailGerenteLoja = contatos.gerenteEmail || contatos.gerenteLoja;
      emailCompliance = contatos.compliance;
      emailDirRH = contatos.diretorRH;
      emailDirOp = contatos.diretorOp;
    }

    let linksEvidenciasDrive = '';
    if (dados.arquivosEvidencias && dados.arquivosEvidencias.length > 0) {
      linksEvidenciasDrive = salvarEvidenciasDrive(dados.arquivosEvidencias, filial);
    }
    const anexoTexto = [dados.anexo || '', linksEvidenciasDrive].filter(x => x).join(' \n');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let apuracaoSheet = ss.getSheetByName('HISTORICO_APURACAO') || ss.insertSheet('HISTORICO_APURACAO');
    
    if (apuracaoSheet.getLastRow() === 0) {
      apuracaoSheet.appendRow([
        'ID Apuracao', 'Data Registro', 'Origem', 'Filial', 'Diretoria', 'Regional', 'Apurador', 'Conclusao', 'Link Doc', 'Nota Humor Antes', 
        'Tratativa', 'Data Recebimento', 'Data Conclusao', 'Resumo Denuncia', 'Diagnostico Apurado', 'Justificativa', 'Enviar Feedback Gerente', 
        'Feedback Texto', 'ID Gerente', 'Nome Gerente', 'Agendar Intervencao', 'Detalhes Intervencao', 
        'ID Denunciado 1', 'Nome Denunciado 1', 'Filial Denunciado 1',
        'ID Denunciado 2', 'Nome Denunciado 2', 'Filial Denunciado 2',
        'ID Denunciado 3', 'Nome Denunciado 3', 'Filial Denunciado 3',
        'Data Ultima Alteracao', 'Email Criador'
      ]);
    }

    const dataRegistro = new Date().toLocaleDateString('pt-BR');
    const nomesDenunciadosArray = dados.denunciados.map(d => d.nome).filter(n => n);
    const stringDenunciadosUnificada = nomesDenunciadosArray.join(', ') || 'Não Informado';

    let idApuracaoDefinitiva = dados.idApuracao || dados.idRascunho || '';
    let rowIndex = -1;
    let docLinkExistente = '';

    if (idApuracaoDefinitiva) {
      const apData = apuracaoSheet.getDataRange().getValues();
      for (let i = 1; i < apData.length; i++) {
        if (apData[i][0] && apData[i][0].toString().trim().toUpperCase() === idApuracaoDefinitiva.trim().toUpperCase()) {
          rowIndex = i + 1;
          docLinkExistente = apData[i][8] ? apData[i][8].toString() : '';
          const criadorOriginal = apData[i][32] ? apData[i][32].toString().toLowerCase().trim() : '';
          
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode editar este registro.' };
          }
          break;
        }
      }
    }

    let docLink = docLinkExistente;

    if (docLinkExistente && docLinkExistente.includes('document/d/')) {
      const match = docLinkExistente.match(/[-\w]{25,}/);
      if (match) {
        try {
          const doc = DocumentApp.openById(match[0]);
          const body = doc.getBody();
          body.setText(''); 

          body.appendParagraph("RELATÓRIO DE APURAÇÃO E DOSSIÊ FINAL").setHeading(DocumentApp.ParagraphHeading.HEADING1);
          body.appendParagraph("1. DADOS DE CONTROLE").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph("Filial: " + filial + " | Diretoria: " + diretoria + " | Regional: " + regional);
          body.appendParagraph("Apurador: " + (dados.apurador || 'GP'));
          body.appendParagraph("Origem: " + (dados.origem || 'Canal'));
          body.appendParagraph("Denunciado(s): " + stringDenunciadosUnificada);
          body.appendParagraph("Data Recebimento: " + (dados.data_recebimento || 'N/A') + " | Data Conclusão: " + (dados.data_finalizacao || 'N/A'));

          body.appendParagraph("\n2. RESUMO DA DENÚNCIA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph(dados.resumo || '');

          body.appendParagraph("\n3. DIAGNÓSTICO APURADO").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph(dados.diagnostico || '');

          body.appendParagraph("\n4. PARECER FINAL E JUSTIFICATIVA").setHeading(DocumentApp.ParagraphHeading.HEADING2);
          body.appendParagraph("Parecer Conclusivo: " + (dados.conclusao || ''));
          body.appendParagraph("Justificativa: " + (dados.justificativa || ''));
          body.appendParagraph("Plano de Ação/Tratativa: " + (dados.tratativa || ''));
          body.appendParagraph("Evidências / Anexos: " + (anexoTexto || 'Sem evidências adicionais no Drive.'));

          doc.saveAndClose();
        } catch(errDoc){}
      }
    } else {
      const templateFile = DriveApp.getFileById(TEMPLATE_DOC_ID);
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const newFileName = "Relatorio de Apuração F." + filial + " - " + stringDenunciadosUnificada + " - " + dataRegistro;
      const copiedFile = templateFile.makeCopy(newFileName, folder);
      const newDocId = copiedFile.getId();

      const doc = DocumentApp.openById(newDocId);
      const body = doc.getBody();

      body.replaceText('{{filial}}', filial);
      body.replaceText('{{diretoria}}', diretoria);
      body.replaceText('{{regional}}', regional);
      body.replaceText('{{denunciados}}', stringDenunciadosUnificada);
      body.replaceText('{{apurador}}', dados.apurador || '');
      body.replaceText('{{data_recebimento}}', dados.data_recebimento || '');
      body.replaceText('{{data_finalizacao}}', dados.data_finalizacao || '');
      body.replaceText('{{resumo}}', dados.resumo || '');
      body.replaceText('{{diagnostico}}', dados.diagnostico || '');
      body.replaceText('{{conclusao}}', dados.conclusao || '');
      body.replaceText('{{justificativa}}', dados.justificativa || '');
      body.replaceText('{{tratativa}}', dados.tratativa || '');
      body.replaceText('{{gerente_id}}', dados.gerente_id || '');
      body.replaceText('{{gerente_nome}}', dados.gerente_nome || '');
      body.replaceText('{{anexo}}', anexoTexto || 'Sem evidências adicionais no Drive.');
      doc.saveAndClose();
      
      docLink = copiedFile.getUrl();

      aplicarPermissoesArquivo(newDocId, emailCoord, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailGerGP, 'LEITOR');
      if (dados.origem === 'Canal') {
        aplicarPermissoesArquivo(newDocId, emailCompliance, 'LEITOR');
        aplicarPermissoesArquivo(newDocId, emailDirRH, 'LEITOR');
      }
      if (emailRegGP) aplicarPermissoesArquivo(newDocId, emailRegGP, 'LEITOR');
    }

    if (!idApuracaoDefinitiva || rowIndex === -1) {
      idApuracaoDefinitiva = 'APU-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idApuracaoDefinitiva, dataRegistro, dados.origem || 'Canal', filial, diretoria, regional, dados.apurador || 'GP', dados.conclusao,
      docLink, dados.nota_humor_antes || '', dados.tratativa, dados.data_recebimento, dados.data_finalizacao, dados.resumo,
      dados.diagnostico, dados.justificativa, dados.enviar_feedback_gerente || 'nao', dados.feedback_gerente,
      dados.gerente_id, dados.gerente_nome, dados.agendar_intervencao || 'nao', dados.detalhes_intervencao,
      dados.denunciados[0]?.id || '', dados.denunciados[0]?.nome || '', dados.denunciados[0]?.filial || '',
      dados.denunciados[1]?.id || '', dados.denunciados[1]?.nome || '', dados.denunciados[1]?.filial || '',
      dados.denunciados[2]?.id || '', dados.denunciados[2]?.nome || '', dados.denunciados[2]?.filial || '',
      new Date().toLocaleString('pt-BR'), emailLogado
    ];

    if (rowIndex !== -1) {
      apuracaoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      apuracaoSheet.appendRow(rowData);
    }

    if (dados.idRascunho) {
      let rascSheet = ss.getSheetByName('Rascunhos_Apuracoes');
      if (rascSheet) {
        const rascVals = rascSheet.getDataRange().getValues();
        for (let i = 1; i < rascVals.length; i++) {
          if (rascVals[i][0].toString().trim().toUpperCase() === dados.idRascunho.toString().trim().toUpperCase()) {
            rascSheet.deleteRow(i + 1);
            break;
          }
        }
      }
    }

    if (dados.enviar_feedback_gerente === 'sim') {
      let fbSheet = ss.getSheetByName('Feedbacks_Gerentes') || ss.insertSheet('Feedbacks_Gerentes');
      if (fbSheet.getLastRow() === 0) {
        fbSheet.appendRow(['ID Feedback', 'Filial', 'Data Envio', 'Feedback Solicitado', 'E-mail Gerente', 'Status', 'Data Resposta', 'Considerações Gerente', 'Links Anexos', 'Nome Gerente', 'ID Gerente']);
      }
      
      const feedbackId = 'FB-' + Utilities.getUuid().substring(0, 8).toUpperCase();
      fbSheet.appendRow([feedbackId, filial, new Date(), dados.feedback_gerente, emailGerenteLoja || 'Não cadastrado', 'Pendente', '', '', '', dados.gerente_nome || '', dados.gerente_id || '']);

      const webAppUrl = ScriptApp.getService().getUrl();
      const feedbackLink = webAppUrl + "?page=gerente&idFeedback=" + feedbackId;
      
      const copiados = [Session.getActiveUser().getEmail()];
      if (emailRegGP) emailRegGP.split(',').forEach(e => { if (e.trim()) copiados.push(e.trim()) });
      const copiasCC = [...new Set(copiados)].join(',');

      if (emailGerenteLoja) enviarEmailGerenteFeedback(emailGerenteLoja, feedbackId, dados.feedback_gerente, feedbackLink, filial, copiasCC);
    }

    return { sucesso: true, link: docLink, id: idApuracaoDefinitiva };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}

function enviarEmailGerenteFeedback(emailGerente, idFeedback, feedbackText, linkAcesso, filial, emailCopia) {
  const assunto = "[AÇÃO DE CLIMA E FEEDBACK] Diretrizes e Plano de Ação - Filial " + filial;
  const htmlBody = 
    '<div style="font-family: Arial, sans-serif; color: #1a1a1a; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">' +
      '<div style="background-color: #FF5A00; padding: 24px; color: white; text-align: center;">' +
        '<div style="background-color: white; color: #FF5A00; height: 40px; width: 40px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-weight: 900; font-size: 20px; margin-bottom: 10px;">M</div>' +
        '<h2 style="margin: 0; font-size: 20px; font-weight: 800; text-transform: uppercase;">Magazine Luiza</h2>' +
        '<p style="margin: 4px 0 0 0; font-size: 12px; font-weight: bold; opacity: 0.9;">PLANO DE DESENVOLVIMENTO E AÇÃO (FEEDBACK)</p>' +
      '</div><div style="padding: 24px; background-color: white;">' +
        '<p style="font-size: 15px;">Prezado(a) <strong>Gerente da Filial ' + filial + '</strong>,</p>' +
        '<p style="font-size: 14px; color: #4A5568;">Como parte das ações estruturadas da área de Gestão de Pessoas (GP), compartilhamos as seguintes diretrizes de clima para ciente e execução na sua unidade:</p>' +
        '<div style="background-color: #FFF5F5; border-left: 4px solid #FF5A00; padding: 18px; border-radius: 8px; margin: 20px 0; color: #C53030; font-style: italic;">"' + feedbackText + '"</div>' +
        '<p style="font-size: 14px; color: #4A5568;">Para formalizar o recebimento destas diretrizes e registrar as ações adotadas em loja, acesse o painel corporativo clicando abaixo.</p>' +
        '<div style="text-align: center; margin: 30px 0;"><a href="' + linkAcesso + '" target="_blank" rel="noopener noreferrer" style="background-color: #FF5A00; color: white; padding: 14px 28px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Visualizar Diretrizes e Dar Ciente</a></div>' +
      '</div></div>';
  
  const emailsTo = emailGerente.split(';').join(',');
  const opcoes = { to: emailsTo, subject: assunto, htmlBody: htmlBody };
  if (emailCopia) opcoes.cc = emailCopia;
  MailApp.sendEmail(opcoes);
}

// =============================================================================
// 7. DESENVOLVIMENTO / FEEDBACK & LINHA DO TEMPO UNIFICADA DE CLIMA
// =============================================================================
function buscarHistoricoCompletoClima(filial) {
  try {
    const filialNorm = normalizarFilialId(filial);
    const eventos = [];

    const dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      for (let i = 1; i < dataApuracao.length; i++) {
        const fVal = dataApuracao[i][3] ? dataApuracao[i][3].toString().trim() : '';
        if (normalizarFilialId(fVal) === filialNorm) {
          let dtStr = dataApuracao[i][1];
          if (dtStr instanceof Date) dtStr = dtStr.toLocaleDateString('pt-BR');
          
          const denunciados = [
            dataApuracao[i][23] ? dataApuracao[i][23].toString() : '',
            dataApuracao[i][26] ? dataApuracao[i][26].toString() : '',
            dataApuracao[i][29] ? dataApuracao[i][29].toString() : ''
          ].filter(n => n).join(', ');

          eventos.push({
            tipo: 'PAI',
            id: dataApuracao[i][0] ? dataApuracao[i][0].toString() : '',
            data: dtStr || 'N/A',
            titulo: 'Relatório de Apuração (PAI)',
            conclusao: dataApuracao[i][7] ? dataApuracao[i][7].toString() : 'Concluído',
            feedbackGerente: dataApuracao[i][17] ? dataApuracao[i][17].toString() : '',
            envolvidos: denunciados || 'Não informados',
            linkDoc: dataApuracao[i][8] ? dataApuracao[i][8].toString() : ''
          });
        }
      }
    }

    const dataInt = getCachedSheetData(SPREADSHEET_ID, 'Intervencoes_Feedback');
    if (dataInt && dataInt.length > 1) {
      for (let i = 1; i < dataInt.length; i++) {
        const fVal = dataInt[i][1] ? dataInt[i][1].toString().trim() : '';
        if (normalizarFilialId(fVal) === filialNorm) {
          let dtStr = dataInt[i][0];
          if (dtStr instanceof Date) dtStr = dtStr.toLocaleDateString('pt-BR');

          const colabs = [
            dataInt[i][8] ? dataInt[i][8].toString() : '',
            dataInt[i][11] ? dataInt[i][11].toString() : '',
            dataInt[i][14] ? dataInt[i][14].toString() : ''
          ].filter(n => n).join(', ');

          eventos.push({
            tipo: 'FEEDBACK',
            id: dataInt[i][0] ? dataInt[i][0].toString() : '',
            data: dtStr || 'N/A',
            titulo: 'Acompanhamento / Feedback de Clima',
            status: dataInt[i][3] ? dataInt[i][3].toString() : 'Registrado',
            resumo: dataInt[i][4] ? dataInt[i][4].toString() : 'N/A',
            envolvidos: colabs || 'Equipe Geral / Loja',
            notaDepois: dataInt[i][5] ? dataInt[i][5].toString() : '',
            linkDoc: dataInt[i][2] ? dataInt[i][2].toString() : ''
          });
        }
      }
    }

    return eventos;
  } catch (e) {
    return [];
  }
}

function buscarCasosRecentesPorFilial(filial) {
  try {
    const filialNorm = normalizarFilialId(filial);
    const resultados = [];

    const dataApuracao = getCachedSheetData(SPREADSHEET_ID, 'HISTORICO_APURACAO');
    if (dataApuracao && dataApuracao.length > 1) {
      const headers = dataApuracao[0].map(h => normalizarTexto(h));
      const colFilial = headers.indexOf('filial') !== -1 ? headers.indexOf('filial') : 3;
      const colDate = headers.indexOf('data registro') !== -1 ? headers.indexOf('data registro') : 1;
      const colConclusao = headers.indexOf('conclusao') !== -1 ? headers.indexOf('conclusao') : 7;
      const colApurador = headers.indexOf('apurador') !== -1 ? headers.indexOf('apurador') : 6;
      const colLink = headers.indexOf('link doc') !== -1 ? headers.indexOf('link doc') : 8;

      for (let i = 1; i < dataApuracao.length; i++) {
        const fVal = dataApuracao[i][colFilial] ? dataApuracao[i][colFilial].toString().trim() : '';
        if (normalizarFilialId(fVal) === filialNorm) {
          const contatos = obterContatosPorFilial(filial);
          let dataStr = '';
          if (dataApuracao[i][colDate] instanceof Date) dataStr = dataApuracao[i][colDate].toLocaleDateString('pt-BR');
          else dataStr = dataApuracao[i][colDate] ? dataApuracao[i][colDate].toString() : 'N/A';

          const d1 = dataApuracao[i][23] ? dataApuracao[i][23].toString() : '';
          const d2 = dataApuracao[i][26] ? dataApuracao[i][26].toString() : '';
          const d3 = dataApuracao[i][29] ? dataApuracao[i][29].toString() : '';
          const consolidadoDenunciados = [d1, d2, d3].filter(n => n).join(', ') || 'N/A';
          const link = dataApuracao[i][colLink] ? dataApuracao[i][colLink].toString() : '';

          resultados.push({
            dataRegistro: dataStr,
            conclusao: dataApuracao[i][colConclusao] ? dataApuracao[i][colConclusao].toString() : 'Pendente',
            apurador: dataApuracao[i][colApurador] ? dataApuracao[i][colApurador].toString() : 'GP',
            denunciados: consolidadoDenunciados,
            linkDoc: link,
            regional: contatos ? contatos.regional : ''
          });
        }
      }
    }

    return resultados;
  } catch (err) {
    Logger.log("Erro ao buscar casos recentes por filial: " + err.message);
    return [];
  }
}

function processarNovaIntervencao(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    
    const filial = normalizarFilialId(dados.filial);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Intervencoes_Feedback') || ss.insertSheet('Intervencoes_Feedback');
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'ID Intervencao', 'Filial', 'Relatório PAI (Link)', 'Status Evolução', 'Detalhes Atividades', 'Nota Humor Depois', 'Nova Data Programada',
        'ID Colaborador 1', 'Nome Colaborador 1', 'Filial Colaborador 1',
        'ID Colaborador 2', 'Nome Colaborador 2', 'Filial Colaborador 2',
        'ID Colaborador 3', 'Nome Colaborador 3', 'Filial Colaborador 3',
        'Email Criador', 'ID Gerente Alvo', 'Nome Gerente Alvo'
      ]);
    }

    let idIntervencao = dados.id || '';
    let rowIndex = -1;

    if (idIntervencao) {
      const vals = sheet.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (vals[i][0] && vals[i][0].toString().trim().toUpperCase() === idIntervencao.trim().toUpperCase()) {
          const criadorOriginal = vals[i][16] ? vals[i][16].toString().toLowerCase().trim() : '';
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode editar este registro.' };
          }
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (!idIntervencao || rowIndex === -1) {
      idIntervencao = 'INT-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idIntervencao, filial, dados.linkDoc || 'Feedback Avulso / Orientativo (Sem PAI)', 
      dados.status_evolucao, dados.detalhes_intervencao, dados.nota_humor_depois || '', dados.data_intervencao || 'N/A',
      dados.colaboradores[0]?.id || '', dados.colaboradores[0]?.nome || '', dados.colaboradores[0]?.filial || '',
      dados.colaboradores[1]?.id || '', dados.colaboradores[1]?.nome || '', dados.colaboradores[1]?.filial || '',
      dados.colaboradores[2]?.id || '', dados.colaboradores[2]?.nome || '', dados.colaboradores[2]?.filial || '',
      emailLogado, dados.gerente_alvo_id || '', dados.gerente_alvo_nome || ''
    ];

    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { sucesso: true, id: idIntervencao };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// 8. DESLIGAMENTO E CASCATA DE APROVAÇÕES
// =============================================================================
function processarNovoDesligamento(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();

    const filial = normalizarFilialId(dados.filial);
    const contatos = obterContatosPorFilial(filial);
    let diretoria = 'MG/CO', regional = 'Brasília', emailRegionalGP = '', emailDiretorRH = '', emailDiretorOp = '';

    if (contatos) {
      diretoria = contatos.diretoria || diretoria;
      regional = contatos.regional || regional;
      emailRegionalGP = contatos.regionalEmail;
      emailDiretorRH = contatos.diretorRH;
      emailDiretorOp = contatos.diretorOp;
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let fullDesligamentoSheet = ss.getSheetByName('HISTORICO_DESLIGAMENTO_F') || ss.insertSheet('HISTORICO_DESLIGAMENTO_F');
    
    if (fullDesligamentoSheet.getLastRow() === 0) {
      fullDesligamentoSheet.appendRow([
        'ID Desligamento', 'Data Registro', 'Filial', 'Colaborador', 'ID', 'Tempo Empresa', 'Cargo', 'Resultados', 'Justificativa', 'Evidencias', 
        'Parecer Coordenador', 'Email Regional', 'Status Regional', 'Parecer Regional', 'Email Diretor Op', 'Status Diretor Op', 'Parecer Diretor Op', 
        'Email Diretor RH', 'Coordenador Nome', 'Links Imagens Resultados', 'Link Apuração', 'Email Criador'
      ]);
    }

    let idDesligamento = dados.id || '';
    let rowIndex = -1;
    let docLinkExistente = '';

    if (idDesligamento) {
      const desData = fullDesligamentoSheet.getDataRange().getValues();
      for (let i = 1; i < desData.length; i++) {
        if (desData[i][0] && desData[i][0].toString().trim().toUpperCase() === idDesligamento.trim().toUpperCase()) {
          const criadorOriginal = desData[i][21] ? desData[i][21].toString().toLowerCase().trim() : '';
          if (criadorOriginal && criadorOriginal !== emailLogado && !verificarEhAdminMaster(emailLogado)) {
            return { sucesso: false, erro: 'Acesso Negado: Apenas o criador original pode editar este registro.' };
          }
          rowIndex = i + 1;
          docLinkExistente = desData[i][20] ? desData[i][20].toString() : '';
          break;
        }
      }
    }

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const linksImagens = [];
    if (dados.arquivos_resultados && dados.arquivos_resultados.length > 0) {
      dados.arquivos_resultados.forEach((arq, index) => {
        if (arq.dados && arq.dados.includes(',')) {
          const splitData = arq.dados.split(',');
          const contentType = splitData[0].match(/:(.*?);/)[1];
          const rawData = Utilities.base64Decode(splitData[1]);
          const fileName = "Resultados_F" + filial + "_" + dados.colaborador_id + "_" + index + "_" + (arq.nome || ".png");
          linksImagens.push(folder.createFile(Utilities.newBlob(rawData, contentType, fileName)).getUrl());
        }
      });
    }

    let docLink = docLinkExistente;

    if (!docLinkExistente) {
      const templateFile = DriveApp.getFileById(TEMPLATE_DESLIGAMENTO_DOC_ID);
      const copiedFile = templateFile.makeCopy("Relatório de Desligamento F." + filial + " - " + dados.colaborador_nome, folder);
      const newDocId = copiedFile.getId();
      const doc = DocumentApp.openById(newDocId);
      const body = doc.getBody();

      body.replaceText('{{filial}}', filial);
      body.replaceText('{{diretoria}}', diretoria);
      body.replaceText('{{regional}}', regional);
      body.replaceText('{{colaborador_nome}}', dados.colaborador_nome);
      body.replaceText('{{colaborador_id}}', dados.colaborador_id); 
      body.replaceText('{{colaborador_cargo}}', dados.colaborador_cargo);
      body.replaceText('{{tempo_empresa}}', dados.tempo_empresa || 'Não informado');
      body.replaceText('{{data_registro}}', new Date().toLocaleDateString('pt-BR'));
      body.replaceText('{{coordenador_nome}}', dados.coordenador_nome);
      body.replaceText('{{resultados}}', dados.resultados);
      body.replaceText('{{resultados_imagens}}', linksImagens.length > 0 ? linksImagens.join(' \n') : 'Nenhuma imagem.');
      body.replaceText('{{justificativa}}', dados.justificativa);
      body.replaceText('{{evidencias}}', dados.evidencias || 'Sem link.');
      body.replaceText('{{parecer_coordenador}}', dados.parecer_coordenador);
      doc.saveAndClose();

      docLink = copiedFile.getUrl();

      aplicarPermissoesArquivo(newDocId, emailLogado, 'EDITOR');
      aplicarPermissoesArquivo(newDocId, emailRegionalGP, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailDiretorOp, 'LEITOR');
      aplicarPermissoesArquivo(newDocId, emailDiretorRH, 'LEITOR');
    }

    if (!idDesligamento || rowIndex === -1) {
      idDesligamento = 'DES-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    }

    const rowData = [
      idDesligamento, new Date(), filial, dados.colaborador_nome, dados.colaborador_id, dados.tempo_empresa, dados.colaborador_cargo,
      dados.resultados, dados.justificativa, dados.evidencias || '', dados.parecer_coordenador, emailRegionalGP, 'Pendente', '',
      emailDiretorOp, 'Pendente', '', emailDiretorRH, dados.coordenador_nome, linksImagens.join(' \n'), docLink, emailLogado
    ];

    if (rowIndex !== -1) {
      fullDesligamentoSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      fullDesligamentoSheet.appendRow(rowData);
    }

    return { sucesso: true, link: docLink, id: idDesligamento };
  } catch (err) {
    return { sucesso: false, erro: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// 9. FEEDBACKS DO GERENTE E RESPOSTAS
// =============================================================================
function buscarDadosFeedbackGerente(idFeedback) {
  try {
    if (!idFeedback) return { erro: 'ID do feedback não fornecido.' };
    const data = getCachedSheetData(SPREADSHEET_ID, 'Feedbacks_Gerentes');
    if (!data) return { erro: 'Base de feedbacks não localizada.' };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === idFeedback.toString().trim().toUpperCase()) {
        return { 
          id: idFeedback, 
          filial: normalizarFilialId(data[i][1]), 
          feedback: data[i][3] ? data[i][3].toString() : '',
          status: data[i][5] ? data[i][5].toString() : 'Pendente'
        };
      }
    }
    return { erro: 'Plano de ação / Feedback ' + idFeedback + ' não localizado.' };
  } catch (err) {
    return { erro: 'Erro ao carregar feedback: ' + err.toString() };
  }
}

function processarRespostaGerente(idFeedback, consideracoes, arquivosBase64) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Feedbacks_Gerentes');
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1, filial = '', feedbackOriginal = '', nomeGerOriginal = '', idGerOriginal = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === idFeedback) { 
        rowIndex = i + 1; 
        filial = normalizarFilialId(data[i][1]); 
        feedbackOriginal = data[i][3];
        nomeGerOriginal = data[i][9] || '';
        idGerOriginal = data[i][10] || '';
        break; 
      }
    }
    if (rowIndex === -1) return { sucesso: false, erro: 'Feedback não localizado.' };

    const emailLogado = Session.getActiveUser().getEmail();
    const nomeLogado = obterUsuarioLogado(emailLogado).nome;
    const nomeFinalGestor = nomeGerOriginal || nomeLogado;
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const linksAnexos = [];
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyyy_HHmm");

    const docTemp = DocumentApp.create("Termo_Ciente_F" + filial + "_" + timeStamp);
    const body = docTemp.getBody();
    
    body.appendParagraph("PLANO DE AÇÃO E TERMO DE CIÊNCIA DE FEEDBACK").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph("Filial: " + filial + " | Gestor: " + nomeFinalGestor);
    body.appendParagraph("Data da Confirmação: " + new Date().toLocaleString('pt-BR'));
    body.appendParagraph("\n1. DIRETRIZES DO GP:\n" + feedbackOriginal);
    body.appendParagraph("\n2. AÇÕES ADOTADAS PELO GESTOR:\n" + consideracoes);
    docTemp.saveAndClose();
    
    const pdfBlob = docTemp.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName("Termo_Ciente_F" + filial + "_" + timeStamp + ".pdf");
    DriveApp.getFileById(docTemp.getId()).setTrashed(true);
    
    linksAnexos.push(pdfFile.getUrl());

    if (arquivosBase64 && arquivosBase64.length > 0) {
      arquivosBase64.forEach(arq => {
        if (arq.dados && arq.dados.includes(',')) {
          const split = arq.dados.split(',');
          const blob = Utilities.newBlob(Utilities.base64Decode(split[1]), split[0].match(/:(.*?);/)[1], "Comprovante_F" + filial + "_" + arq.nome);
          linksAnexos.push(folder.createFile(blob).getUrl());
        }
      });
    }

    sheet.getRange(rowIndex, 6).setValue('Respondido');
    sheet.getRange(rowIndex, 7).setValue(new Date());
    sheet.getRange(rowIndex, 8).setValue(consideracoes);
    sheet.getRange(rowIndex, 9).setValue(linksAnexos.join(' \n'));

    return { sucesso: true };
  } catch (err) { 
    return { sucesso: false, erro: err.toString() }; 
  } finally {
    lock.releaseLock();
  }
}
