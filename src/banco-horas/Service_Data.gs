/**
 * SISTEMA AUTOMATIZADO DE BANCO DE HORAS
 * Arquivo: Service_Data.gs (Leitura e cruzamento de bases de dados)
 */

/**
 * Mapeia o ID do colaborador para a sua Situação a partir da aba DADOS_ATIVOS (se existir)
 */
function carregarStatusColaboradores(ss) {
  let statusMap = {};
  const abaDadosAtivos = ss.getSheetByName(CONFIG.nomeAbaDadosAtivos);
  if (abaDadosAtivos) {
    const dadosAtivos = abaDadosAtivos.getDataRange().getValues();
    if (dadosAtivos.length > 1) {
      const h = dadosAtivos[0];
      const cId = obterIndice(h, ['id', 'matricula', 'chapa'], 0);
      const cSit = obterIndice(h, ['situacao', 'status'], 6);
      for (let i = 1; i < dadosAtivos.length; i++) {
        let id = dadosAtivos[i][cId] !== undefined && dadosAtivos[i][cId] !== null ? dadosAtivos[i][cId].toString().trim() : "";
        let situacao = dadosAtivos[i][cSit] !== undefined && dadosAtivos[i][cSit] !== null ? dadosAtivos[i][cSit].toString().trim() : "";
        if (id) {
          statusMap[id] = situacao;
        }
      }
    }
  }
  return statusMap;
}

function getRegionaisPortal() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.nomeAbaFiliais);
    if (!aba) return ["⚠️ ERRO: Aba '" + CONFIG.nomeAbaFiliais + "' não existe!"];
    
    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return ["⚠️ ERRO: Aba '" + CONFIG.nomeAbaFiliais + "' está vazia!"];
    
    let regionais = [...new Set(dados.slice(1).map(r => r[1]).filter(String))];
    let diretorias = [...new Set(dados.slice(1).map(r => r[2]).filter(String))];
    
    let listaFinal = [...new Set([...regionais, ...diretorias])];
    if (listaFinal.length === 0) return ["⚠️ ERRO: Nenhuma regional encontrada"];
    
    return listaFinal.sort();
  } catch (e) {
    return ["⚠️️ ERRO FATAL: " + e.message];
  }
}

function getColaboradoresPortal(tipo, parametro) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const aba = ss.getSheetByName(CONFIG.nomeAbaBase);
    const abaFiliais = ss.getSheetByName(CONFIG.nomeAbaFiliais);
    const abaFolgas = ss.getSheetByName(CONFIG.nomeAbaFolgas);
    
    if (!aba) throw new Error("Aba '" + CONFIG.nomeAbaBase + "' não encontrada.");
    
    const statusMap = carregarStatusColaboradores(ss);
    
    let mapaFilialDiretoria = {};
    let mapaFilialRegional = {};
    if (abaFiliais) {
      let dadosF = abaFiliais.getDataRange().getValues();
      for (let f = 1; f < dadosF.length; f++) {
        let fId = normalizarFilialId(dadosF[f][0]);
        mapaFilialRegional[fId] = padronizarTexto(dadosF[f][1]); // Coluna B
        mapaFilialDiretoria[fId] = padronizarTexto(dadosF[f][2]); // Coluna C
      }
    }
    
    // Mapear folgas cadastradas mais recentes por ID do colaborador
    let folgasMap = {};
    if (abaFolgas) {
      const dadosFolgas = abaFolgas.getDataRange().getValues();
      for (let i = 1; i < dadosFolgas.length; i++) {
        let idColab = dadosFolgas[i][3] ? dadosFolgas[i][3].toString().trim() : "";
        let dataFolga = dadosFolgas[i][4];
        
        if (dataFolga instanceof Date) {
          folgasMap[idColab] = Utilities.formatDate(dataFolga, "GMT-3", "dd/MM/yyyy");
        } else if (dataFolga) {
          let strData = dataFolga.toString().trim();
          let parts = strData.split('-');
          if (parts.length === 3) {
            folgasMap[idColab] = `${parts[2]}/${parts[1]}/${parts[0]}`;
          } else {
            folgasMap[idColab] = strData;
          }
        }
      }
    }
    
    const dadosBase = aba.getDataRange().getValues();
    if (dadosBase.length <= 1) return [];

    // Detecção dinâmica das colunas com base no layout real da aba
    const hBase = dadosBase[0];
    const cFilial   = obterIndice(hBase, ['filial'], 0);
    const cDir      = obterIndice(hBase, ['diretoria'], 2);
    const cReg      = obterIndice(hBase, ['regional'], 3);
    const cId       = obterIndice(hBase, ['id', 'matricula', 'chapa'], 4);
    const cNome     = obterIndice(hBase, ['nome', 'colaborador'], 5);
    const cCargo    = obterIndice(hBase, ['cargo'], 6);
    const cAtual    = obterIndice(hBase, ['atual'], 8);

    let resultados = [];
    let paramLimpo = padronizarTexto(parametro);
    let filialParamNorm = normalizarFilialId(parametro);

    for (let i = 1; i < dadosBase.length; i++) {
      let row = dadosBase[i];
      
      let filial        = normalizarFilialId(row[cFilial]);
      let id            = row[cId] ? row[cId].toString().trim() : "";
      let nome          = row[cNome] ? row[cNome].toString().trim() : "";
      let regionalOrig  = row[cReg] ? row[cReg].toString().trim() : "";
      let regional      = padronizarTexto(regionalOrig);
      let diretoria     = row[cDir] ? padronizarTexto(row[cDir]) : "";
      let cargo         = row[cCargo] ? row[cCargo].toString().trim() : "";
      let horas         = parseHoras(row[cAtual]); // Saldo Atual Vigente
      
      let situacaoFinal = statusMap[id] || "";
      if (!isEmAtividadeNormal(situacaoFinal)) {
        continue;
      }

      if (horas >= 5) {
        let match = false;
        
        if (tipo === 'gerente' && filial === filialParamNorm) {
          match = true;
        } else if (tipo === 'coord') {
          let regFilial = mapaFilialRegional[filial] || "";
          let dirFilial = mapaFilialDiretoria[filial] || "";
          
          if (regional === paramLimpo || regional.includes(paramLimpo) || paramLimpo.includes(regional) ||
              diretoria === paramLimpo || diretoria.includes(paramLimpo) || paramLimpo.includes(diretoria) ||
              regFilial === paramLimpo || regFilial.includes(paramLimpo) || paramLimpo.includes(regFilial) ||
              dirFilial === paramLimpo || dirFilial.includes(paramLimpo) || paramLimpo.includes(dirFilial)) {
            match = true;
          }
        }

        if (match) {
          resultados.push({ 
            filial: filial, 
            regional: regionalOrig, 
            id: id, 
            nome: nome, 
            cargo: cargo,
            horas: horas,
            dataAgendada: folgasMap[id] || null
          });
        }
      }
    }
    return resultados;
  } catch (e) {
    return [{ filial: 'ERRO', regional: 'ERRO', id: '000', nome: 'Falha: ' + e.message, horas: 0, dataAgendada: null }];
  }
}
