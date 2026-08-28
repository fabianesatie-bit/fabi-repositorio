/**
 * SERVIÇO DE DADOS - CONSOLIDAÇÃO DA AUDITORIA COM DADOS_LOJAS
 */

function obterDadosAuditoria() {
  try {
    var perfil = obterPerfilAcessoLogado();
    var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
    var ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);

    // 1. CARREGA E MAPEIA A ABA DADOS_LOJAS
    var shLojas = ssMaster.getSheetByName(TAB_NAMES.LOJAS);
    var mapaLojas = {}; // chave: filialId normalizado -> objeto loja
    var mapaHierarquia = {}; // estrutura para popular os filtros (Diretorias -> Regionais -> Filiais)

    if (shLojas) {
      var dLojas = shLojas.getDataRange().getValues();
      // Assume cabeçalho na linha 0: Filial_ID, Nome_Fantasia, Regional, Diretoria, Gerente_Email, Cidade, Estado
      for (var l = 1; l < dLojas.length; l++) {
        var rawId = dLojas[l][0];
        if (!rawId && rawId !== 0) continue;

        var fidNorm = normalizarFilialId(rawId);
        var nomeFantasia = String(dLojas[l][1] || '').trim();
        var reg = String(dLojas[l][2] || 'NÃO INFORMADA').trim().toUpperCase();
        var dir = String(dLojas[l][3] || 'OUTROS').trim().toUpperCase();
        var cidade = String(dLojas[l][5] || '').trim();
        var estado = String(dLojas[l][6] || '').trim();

        var objLoja = {
          filialId: fidNorm,
          nomeLoja: 'Filial ' + fidNorm + (nomeFantasia ? ' - ' + nomeFantasia : ''),
          regional: reg,
          diretoria: dir,
          cidade: cidade,
          estado: estado
        };

        mapaLojas[fidNorm] = objLoja;

        if (!mapaHierarquia[dir]) {
          mapaHierarquia[dir] = {};
        }
        if (!mapaHierarquia[dir][reg]) {
          mapaHierarquia[dir][reg] = [];
        }
        mapaHierarquia[dir][reg].push(fidNorm);
      }
    }

    // 2. LOCALIZA AS ABAS DE APONTAMENTO
    var sheets = ssAuditoria.getSheets();
    var shForaJornada = null;
    var shHorasExtras = null;
    var shBritanicos = null;

    sheets.forEach(function(sh) {
      var nomeNorm = normalizarTextoUpper(sh.getName());
      if (nomeNorm.includes(normalizarTextoUpper(TAB_NAMES.FORA_JORNADA_PATTERN))) shForaJornada = sh;
      if (nomeNorm.includes(normalizarTextoUpper(TAB_NAMES.HORAS_EXTRAS_PATTERN))) shHorasExtras = sh;
      if (nomeNorm.includes(normalizarTextoUpper(TAB_NAMES.BRITANICOS_PATTERN))) shBritanicos = sh;
    });

    var resumoFiliais = {};
    var colaboradoresSet = new Set();
    var apontamentosCriticosCount = 0;
    var listaApontamentosDetalhada = [];

    // Helper interno para obter metadata de loja
    function getLojaMeta(filialRaw) {
      var fid = normalizarFilialId(filialRaw);
      if (mapaLojas[fid]) return mapaLojas[fid];
      return {
        filialId: fid,
        nomeLoja: 'Filial ' + fid,
        regional: 'OUTRAS',
        diretoria: 'GERAL',
        cidade: '',
        estado: ''
      };
    }

    // Helper interno para registrar filial
    function obterOuCriarResumoFilial(meta) {
      var fid = meta.filialId;
      if (!resumoFiliais[fid]) {
        resumoFiliais[fid] = {
          filial: fid,
          nomeLoja: meta.nomeLoja,
          regional: meta.regional,
          diretoria: meta.diretoria,
          cidade: meta.cidade,
          estado: meta.estado,
          totalApontamentos: 0,
          foraJornadaCount: 0,
          horasExtrasCount: 0,
          britanicosCount: 0,
          colaboradores: new Set(),
          risco: 'Baixo Risco'
        };
      }
      return resumoFiliais[fid];
    }

    // 3. PROCESSA ABRANGÊNCIA: Acesso fora da jornada
    if (shForaJornada) {
      var dFora = shForaJornada.getDataRange().getValues();
      for (var i = 1; i < dFora.length; i++) {
        var filialRaw = dFora[i][2];
        if (!filialRaw && filialRaw !== 0) continue;

        var metaLoja = getLojaMeta(filialRaw);

        // Trava RLS por regional/diretoria
        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLoja.regional)) continue;

        var cdi = String(dFora[i][4] || '').trim();
        var nome = String(dFora[i][5] || '').trim();
        var keyColab = cdi || nome;

        if (keyColab) colaboradoresSet.add(keyColab);
        apontamentosCriticosCount++;

        var rf = obterOuCriarResumoFilial(metaLoja);
        rf.totalApontamentos++;
        rf.foraJornadaCount++;
        if (keyColab) rf.colaboradores.add(keyColab);

        listaApontamentosDetalhada.push({
          tipo: 'FORA_JORNADA',
          filial: metaLoja.filialId,
          regional: metaLoja.regional,
          diretoria: metaLoja.diretoria,
          colaborador: keyColab
        });
      }
    }

    // 4. PROCESSA ABRANGÊNCIA: Horas extras (>2h)
    if (shHorasExtras) {
      var dHE = shHorasExtras.getDataRange().getValues();
      for (var j = 1; j < dHE.length; j++) {
        var filialRawHE = dHE[j][0];
        if (!filialRawHE && filialRawHE !== 0) continue;

        var metaLojaHE = getLojaMeta(filialRawHE);

        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLojaHE.regional)) continue;

        var chapa = String(dHE[j][1] || '').trim();
        var nomeHE = String(dHE[j][2] || '').trim();
        var keyColabHE = chapa || nomeHE;

        if (keyColabHE) colaboradoresSet.add(keyColabHE);

        var rfHE = obterOuCriarResumoFilial(metaLojaHE);
        rfHE.totalApontamentos++;
        rfHE.horasExtrasCount++;
        if (keyColabHE) rfHE.colaboradores.add(keyColabHE);

        listaApontamentosDetalhada.push({
          tipo: 'HORAS_EXTRAS',
          filial: metaLojaHE.filialId,
          regional: metaLojaHE.regional,
          diretoria: metaLojaHE.diretoria,
          colaborador: keyColabHE
        });
      }
    }

    // 5. PROCESSA ABRANGÊNCIA: Ajustes Britânicos
    if (shBritanicos) {
      var dBrit = shBritanicos.getDataRange().getValues();
      for (var k = 1; k < dBrit.length; k++) {
        var filialRawB = dBrit[k][6];
        if (!filialRawB && filialRawB !== 0) continue;

        var metaLojaB = getLojaMeta(filialRawB);
        var qtdBritanicos = parseFloatBR(dBrit[k][10]);
        if (qtdBritanicos <= 0) continue;

        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLojaB.regional)) continue;

        var rfB = obterOuCriarResumoFilial(metaLojaB);
        rfB.totalApontamentos += qtdBritanicos;
        rfB.britanicosCount += qtdBritanicos;

        listaApontamentosDetalhada.push({
          tipo: 'BRITANICO',
          filial: metaLojaB.filialId,
          regional: metaLojaB.regional,
          diretoria: metaLojaB.diretoria,
          quantidade: qtdBritanicos
        });
      }
    }

    // Classificação de Risco por Filial
    var listaFiliais = Object.values(resumoFiliais).map(function(f) {
      if (f.totalApontamentos >= 15) {
        f.risco = 'Alto Risco';
      } else if (f.totalApontamentos >= 5) {
        f.risco = 'Médio Risco';
      } else {
        f.risco = 'Baixo Risco';
      }
      f.qtdColaboradores = f.colaboradores.size;
      delete f.colaboradores; // remove o Set antes da serialização JSON
      return f;
    });

    listaFiliais.sort(function(a, b) { return b.totalApontamentos - a.totalApontamentos; });

    return JSON.stringify({
      sucesso: true,
      perfil: perfil,
      mapaHierarquia: mapaHierarquia,
      filiaisAlertas: listaFiliais,
      apontamentosBrutos: listaApontamentosDetalhada,
      kpisGlobais: {
        colaboradoresIrregulares: colaboradoresSet.size,
        apontamentosCriticos: apontamentosCriticosCount
      }
    });

  } catch (e) {
    return JSON.stringify({ sucesso: false, erro: e.message || e.toString() });
  }
}
