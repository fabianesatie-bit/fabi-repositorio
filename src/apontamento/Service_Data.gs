/**
 * SERVIÇO DE DADOS E CONSOLIDAÇÃO DA AUDITORIA
 */

function obterDadosAuditoria() {
  try {
    var perfil = obterPerfilAcessoLogado();
    var ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    var sheets = ssAuditoria.getSheets();

    var shForaJornada = null;
    var shHorasExtras = null;
    var shBritanicos = null;

    // Busca dinâmica de abas por aproximação de nome
    sheets.forEach(function(sh) {
      var nomeNorm = normalizarTexto(sh.getName());
      if (nomeNorm.includes(normalizarTexto(TAB_NAMES.FORA_JORNADA_PATTERN))) shForaJornada = sh;
      if (nomeNorm.includes(normalizarTexto(TAB_NAMES.HORAS_EXTRAS_PATTERN))) shHorasExtras = sh;
      if (nomeNorm.includes(normalizarTexto(TAB_NAMES.BRITANICOS_PATTERN))) shBritanicos = sh;
    });

    var resumoFiliais = {};
    var regionaisRisco = {};
    var colaboradoresSet = new Set();
    var apontamentosCriticosCount = 0;

    // 1. Processa Acesso Fora da Jornada
    if (shForaJornada) {
      var dFora = shForaJornada.getDataRange().getValues();
      for (var i = 1; i < dFora.length; i++) {
        var reg = String(dFora[i][0] || 'NÃO INFORMADA').toUpperCase().trim();
        var dir = String(dFora[i][1] || 'GERAL').toUpperCase().trim();
        var filialRaw = dFora[i][2];
        var filialId = normalizarFilialId(filialRaw);
        var cidade = String(dFora[i][3] || '').trim();
        var cdi = String(dFora[i][4] || '').trim();
        var nome = String(dFora[i][5] || '').trim();
        var cargo = String(dFora[i][6] || '').trim();
        var irregularidade = String(dFora[i][7] || 'Fora da Jornada').trim();

        if (!filialId) continue;

        // Trava RLS
        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(reg)) continue;

        colaboradoresSet.add(cdi || nome);
        apontamentosCriticosCount++;

        if (!resumoFiliais[filialId]) {
          resumoFiliais[filialId] = {
            filial: filialId,
            nomeLoja: 'Filial ' + filialId + (cidade ? ' - ' + cidade : ''),
            regional: reg,
            diretoria: dir,
            cidade: cidade,
            totalApontamentos: 0,
            foraJornadaCount: 0,
            horasExtrasCount: 0,
            britanicosCount: 0,
            risco: 'Baixo Risco',
            colaboradores: {}
          };
        }

        resumoFiliais[filialId].totalApontamentos++;
        resumoFiliais[filialId].foraJornadaCount++;

        if (!regionaisRisco[reg]) regionaisRisco[reg] = { regional: reg, quantidade: 0 };
        regionaisRisco[reg].quantidade++;
      }
    }

    // 2. Processa Horas Extras (> 2h)
    if (shHorasExtras) {
      var dHE = shHorasExtras.getDataRange().getValues();
      for (var j = 1; j < dHE.length; j++) {
        var filialIdHE = normalizarFilialId(dHE[j][0]);
        var chapa = String(dHE[j][1] || '').trim();
        var nomeHE = String(dHE[j][2] || '').trim();
        var cargoHE = String(dHE[j][3] || '').trim();
        var qtdHorasStr = String(dHE[j][8] || '').trim();

        if (!filialIdHE) continue;

        var filialRef = resumoFiliais[filialIdHE];
        var regHE = filialRef ? filialRef.regional : 'OUTRAS';

        if (!perfil.temAcessoTotal && filialRef && !perfil.regionaisPermitidas.includes(regHE)) continue;

        colaboradoresSet.add(chapa || nomeHE);

        if (!resumoFiliais[filialIdHE]) {
          resumoFiliais[filialIdHE] = {
            filial: filialIdHE,
            nomeLoja: 'Filial ' + filialIdHE,
            regional: regHE,
            diretoria: 'GERAL',
            cidade: '',
            totalApontamentos: 0,
            foraJornadaCount: 0,
            horasExtrasCount: 0,
            britanicosCount: 0,
            risco: 'Baixo Risco',
            colaboradores: {}
          };
        }

        resumoFiliais[filialIdHE].totalApontamentos++;
        resumoFiliais[filialIdHE].horasExtrasCount++;
      }
    }

    // 3. Processa Ajustes Britânicos
    if (shBritanicos) {
      var dBrit = shBritanicos.getDataRange().getValues();
      for (var k = 1; k < dBrit.length; k++) {
        var regB = String(dBrit[k][0] || '').toUpperCase().trim();
        var dirB = String(dBrit[k][4] || '').toUpperCase().trim();
        var filialIdB = normalizarFilialId(dBrit[k][6]);
        var qtdBritanicos = parseFloatBR(dBrit[k][10]);

        if (!filialIdB) continue;
        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(regB)) continue;

        if (resumoFiliais[filialIdB]) {
          resumoFiliais[filialIdB].britanicosCount += qtdBritanicos;
          resumoFiliais[filialIdB].totalApontamentos += qtdBritanicos;
        }
      }
    }

    // Classificação de Risco por Filial
    var listaFiliais = Object.values(resumoFiliais);
    listaFiliais.forEach(function(f) {
      if (f.totalApontamentos >= 10) {
        f.risco = 'Alto Risco';
      } else if (f.totalApontamentos >= 5) {
        f.risco = 'Médio Risco';
      } else {
        f.risco = 'Baixo Risco';
      }
    });

    listaFiliais.sort(function(a, b) { return b.totalApontamentos - a.totalApontamentos; });

    // Compilação do Ranking de Regionais
    var listaRegionais = Object.values(regionaisRisco);
    listaRegionais.sort(function(a, b) { return b.quantidade - a.quantidade; });

    // Extração de Opções de Filtros
    var diretoriasSet = new Set();
    var regionaisSet = new Set();
    var filiaisSet = new Set();

    listaFiliais.forEach(function(f) {
      if (f.diretoria) diretoriasSet.add(f.diretoria);
      if (f.regional) regionaisSet.add(f.regional);
      if (f.filial) filiaisSet.add(f.filial);
    });

    return JSON.stringify({
      sucesso: true,
      perfil: perfil,
      kpis: {
        colaboradoresIrregulares: colaboradoresSet.size,
        apontamentosCriticos: apontamentosCriticosCount
      },
      topRegionais: listaRegionais.slice(0, 5),
      filiaisAlertas: listaFiliais,
      filtros: {
        diretorias: Array.from(diretoriasSet).sort(),
        regionais: Array.from(regionaisSet).sort(),
        filiais: Array.from(filiaisSet).sort(function(a,b){ return parseInt(a)-parseInt(b); })
      }
    });

  } catch (e) {
    return JSON.stringify({ sucesso: false, erro: e.message || e.toString() });
  }
}

