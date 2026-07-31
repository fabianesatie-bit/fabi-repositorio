// =============================================================================
// FUNÇÕES UTILITÁRIAS E CONSULTAS DE BANCO DE DADOS (DADOS_LOJAS / DADOS_ATIVOS)
// Subpasta GitHub: src/forms-denuncia/
// Arquivo Apps Script: DatabaseUtils.gs
// =============================================================================

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

function normalizarTexto(texto) {
  if (!texto) return '';
  return texto.toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function getCachedSheetData(spreadsheetId, sheetName) {
  const cacheKey = spreadsheetId + '_' + sheetName;
  if (typeof SpreadsheetCache !== 'undefined' && SpreadsheetCache[cacheKey]) {
    return SpreadsheetCache[cacheKey];
  }

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
    if (typeof SpreadsheetCache !== 'undefined') {
      SpreadsheetCache[cacheKey] = data;
    }
    return data;
  } catch (err) {
    Logger.log("Erro ao abrir planilha/aba (" + sheetName + "): " + err.message);
    return null;
  }
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
