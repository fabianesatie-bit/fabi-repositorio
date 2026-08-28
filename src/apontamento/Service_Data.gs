/**
 * SERVIÇO DE DADOS - CONSOLIDAÇÃO DA AUDITORIA E CERTIFICADOS
 */

function obterDadosAuditoria() {
  try {
    var perfil = obterPerfilAcessoLogado();
    var ssMaster = SpreadsheetApp.openById(SPREADSHEET_DB_MASTER_ID);
    var ssAuditoria = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);

    var shLojas = ssMaster.getSheetByName(TAB_NAMES.LOJAS);
    var mapaLojas = {};
    var mapaHierarquia = {};

    if (shLojas) {
      var dLojas = shLojas.getDataRange().getValues();
      for (var l = 1; l < dLojas.length; l++) {
        var rawId = dLojas[l][0];
        if (!rawId && rawId !== 0) continue;

        var fidNorm = normalizarFilialId(rawId);
        var nomeFantasia = String(dLojas[l][1] || '').trim();
        var reg = String(dLojas[l][2] || 'NÃO INFORMADA').trim().toUpperCase();
        var dir = String(dLojas[l][3] || 'OUTROS').trim().toUpperCase();
        var gerenteEmail = String(dLojas[l][4] || '').trim();
        var cidade = String(dLojas[l][5] || '').trim();
        var estado = String(dLojas[l][6] || '').trim();

        var objLoja = {
          filialId: fidNorm,
          nomeLoja: 'Filial ' + fidNorm + (nomeFantasia ? ' - ' + nomeFantasia : ''),
          regional: reg,
          diretoria: dir,
          gerenteEmail: gerenteEmail || ('gerente.filial' + fidNorm + '@magazineluiza.com.br'),
          cidade: cidade,
          estado: estado
        };

        mapaLojas[fidNorm] = objLoja;
        if (!mapaHierarquia[dir]) mapaHierarquia[dir] = {};
        if (!mapaHierarquia[dir][reg]) mapaHierarquia[dir][reg] = [];
        mapaHierarquia[dir][reg].push(fidNorm);
      }
    }

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

    function getLojaMeta(filialRaw) {
      var fid = normalizarFilialId(filialRaw);
      if (mapaLojas[fid]) return mapaLojas[fid];
      return {
        filialId: fid,
        nomeLoja: 'Filial ' + fid,
        regional: 'OUTRAS',
        diretoria: 'GERAL',
        gerenteEmail: 'gplojas@magazineluiza.com.br',
        cidade: '',
        estado: ''
      };
    }

    function obterOuCriarResumoFilial(meta) {
      var fid = meta.filialId;
      if (!resumoFiliais[fid]) {
        resumoFiliais[fid] = {
          filial: fid,
          nomeLoja: meta.nomeLoja,
          regional: meta.regional,
          diretoria: meta.diretoria,
          gerenteEmail: meta.gerenteEmail,
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

    if (shForaJornada) {
      var dFora = shForaJornada.getDataRange().getValues();
      for (var i = 1; i < dFora.length; i++) {
        var filialRaw = dFora[i][2];
        if (!filialRaw && filialRaw !== 0) continue;

        var metaLoja = getLojaMeta(filialRaw);
        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLoja.regional)) continue;

        var cdi = String(dFora[i][4] || '').trim();
        var nome = String(dFora[i][5] || '').trim();
        var cargo = String(dFora[i][6] || '').trim();
        var irr = String(dFora[i][7] || 'Acesso Fora da Jornada').trim();
        var keyColab = cdi || nome;

        if (keyColab) colaboradoresSet.add(keyColab);
        apontamentosCriticosCount++;

        var rf = obterOuCriarResumoFilial(metaLoja);
        rf.totalApontamentos++;
        rf.foraJornadaCount++;
        if (keyColab) rf.colaboradores.add(keyColab);

        listaApontamentosDetalhada.push({
          id: 'AFJ-' + i,
          filialId: metaLoja.filialId,
          filialNome: metaLoja.nomeLoja,
          regional: metaLoja.regional,
          diretoria: metaLoja.diretoria,
          chapa: cdi,
          nome: nome,
          cargo: cargo || 'Assistente de Vendas',
          tipoIrregularidade: irr,
          subtipo: 'Acesso ao PDV/Sistema',
          quantidadeMes: 1
        });
      }
    }

    if (shHorasExtras) {
      var dHE = shHorasExtras.getDataRange().getValues();
      for (var j = 1; j < dHE.length; j++) {
        var filialRawHE = dHE[j][0];
        if (!filialRawHE && filialRawHE !== 0) continue;

        var metaLojaHE = getLojaMeta(filialRawHE);
        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLojaHE.regional)) continue;

        var chapaHE = String(dHE[j][1] || '').trim();
        var nomeHE = String(dHE[j][2] || '').trim();
        var cargoHE = String(dHE[j][3] || '').trim();
        var qtdHoras = String(dHE[j][8] || '2h+').trim();
        var keyColabHE = chapaHE || nomeHE;

        if (keyColabHE) colaboradoresSet.add(keyColabHE);

        var rfHE = obterOuCriarResumoFilial(metaLojaHE);
        rfHE.totalApontamentos++;
        rfHE.horasExtrasCount++;
        if (keyColabHE) rfHE.colaboradores.add(keyColabHE);

        listaApontamentosDetalhada.push({
          id: 'HE-' + j,
          filialId: metaLojaHE.filialId,
          filialNome: metaLojaHE.nomeLoja,
          regional: metaLojaHE.regional,
          diretoria: metaLojaHE.diretoria,
          chapa: chapaHE,
          nome: nomeHE,
          cargo: cargoHE || 'Operador de Loja',
          tipoIrregularidade: 'Mais De 2 Horas Extras Diárias',
          subtipo: 'Excesso de Jornada (' + qtdHoras + ')',
          quantidadeMes: 1
        });
      }
    }

    if (shBritanicos) {
      var dBrit = shBritanicos.getDataRange().getValues();
      for (var k = 1; k < dBrit.length; k++) {
        var filialRawB = dBrit[k][6];
        if (!filialRawB && filialRawB !== 0) continue;

        var metaLojaB = getLojaMeta(filialRawB);
        var qtdBritanicos = parseFloatBR(dBrit[k][10]);
        if (qtdBritanicos <= 0) continue;

        if (!perfil.temAcessoTotal && !perfil.regionaisPermitidas.includes(metaLojaB.regional)) continue;

        var cdiB = String(dBrit[k][5] || '').trim();
        var nomeB = String(dBrit[k][7] || '').trim();
        var cargoB = String(dBrit[k][8] || '').trim();
        var percB = String(dBrit[k][11] || '100%').trim();

        var rfB = obterOuCriarResumoFilial(metaLojaB);
        rfB.totalApontamentos += qtdBritanicos;
        rfB.britanicosCount += qtdBritanicos;

        listaApontamentosDetalhada.push({
          id: 'BRIT-' + k,
          filialId: metaLojaB.filialId,
          filialNome: metaLojaB.nomeLoja,
          regional: metaLojaB.regional,
          diretoria: metaLojaB.diretoria,
          chapa: cdiB,
          nome: nomeB,
          cargo: cargoB || 'Vendedor',
          tipoIrregularidade: 'Ajuste / Marcação Britânica',
          subtipo: 'Marcação Invariável (' + percB + ')',
          quantidadeMes: qtdBritanicos
        });
      }
    }

    var listaFiliais = Object.values(resumoFiliais).map(function(f) {
      if (f.totalApontamentos >= 15) {
        f.risco = 'Alto Risco';
      } else if (f.totalApontamentos >= 5) {
        f.risco = 'Médio Risco';
      } else {
        f.risco = 'Baixo Risco';
      }
      f.qtdColaboradores = f.colaboradores.size;
      delete f.colaboradores;
      return f;
    });

    listaFiliais.sort(function(a, b) { return b.totalApontamentos - a.totalApontamentos; });

    var certificados = [];
    var shCert = ssAuditoria.getSheetByName(TAB_NAMES.CERTIFICADOS);
    if (shCert) {
      var dCert = shCert.getDataRange().getValues();
      for (var c = 1; c < dCert.length; c++) {
        if (!dCert[c][0]) continue;
        certificados.push({
          id: String(dCert[c][0]),
          dataEnvio: String(dCert[c][1]),
          filialId: String(dCert[c][2]),
          filialNome: String(dCert[c][3]),
          chapa: String(dCert[c][4]),
          colaboradorNome: String(dCert[c][5]),
          cargo: String(dCert[c][6]),
          cursoNome: String(dCert[c][7]),
          linkComprovante: String(dCert[c][8]),
          status: String(dCert[c][9] || 'Aprovado'),
          registradoPor: String(dCert[c][10] || ''),
          observacoes: String(dCert[c][11] || '')
        });
      }
    }

    return JSON.stringify({
      sucesso: true,
      perfil: perfil,
      mapaHierarquia: mapaHierarquia,
      filiaisAlertas: listaFiliais,
      apontamentosBrutos: listaApontamentosDetalhada,
      certificados: certificados,
      kpisGlobais: {
        colaboradoresIrregulares: colaboradoresSet.size,
        apontamentosCriticos: apontamentosCriticosCount
      }
    });

  } catch (e) {
    return JSON.stringify({ sucesso: false, erro: e.message || e.toString() });
  }
}

function saveCertificadoTreinamento(certificadoData) {
  try {
    if (!certificadoData) throw new Error('Dados do certificado inválidos.');

    var ss = SpreadsheetApp.openById(SPREADSHEET_AUDITORIA_ID);
    var sheetName = TAB_NAMES.CERTIFICADOS;
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = [
        'ID',
        'Data_Registro',
        'Filial_ID',
        'Filial_Nome',
        'Chapa',
        'Colaborador_Nome',
        'Cargo',
        'Tipo_Treinamento',
        'Link_Certificado',
        'Status',
        'Registrado_Por',
        'Observacoes'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground('#0086ff')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    var dataAgora = Utilities.formatDate(new Date(), 'GMT-3', 'yyyy-MM-dd HH:mm:ss');
    var idNovo = 'CERT-' + new Date().getTime();

    var row = [
      idNovo,
      dataAgora,
      String(certificadoData.filialId || ''),
      String(certificadoData.filialNome || ''),
      String(certificadoData.chapa || ''),
      String(certificadoData.colaboradorNome || ''),
      String(certificadoData.cargo || ''),
      String(certificadoData.tipoTreinamento || 'Ponto Eletrônico - Guia Essencial do Colaborador'),
      String(certificadoData.linkComprovante || ''),
      String(certificadoData.status || 'Aprovado'),
      String(certificadoData.registradoPor || 'Gerente de Loja'),
      String(certificadoData.observacoes || '')
    ];

    sheet.appendRow(row);

    return {
      success: true,
      id: idNovo,
      message: 'Comprovante do colaborador (ID: ' + (certificadoData.chapa || '-') + ') salvo com sucesso na planilha!'
    };
  } catch (err) {
    Logger.log('Erro ao salvar certificado: ' + err.toString());
    return { success: false, error: err.message };
  }
}
