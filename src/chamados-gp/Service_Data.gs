/**
 * ============================================================================
 * MÓDULO: Service_Data.gs
 * PROJETO: Ecossistema GP360 - Atendimento e Chamados GP
 * AUTOR: Desenvolvedor Senior / Arquiteto GP360
 * REPOSITÓRIO: fabianesatie-bit/fabi_repositorio (src/chamados-gp/Service_Data.gs)
 * ============================================================================
 */

/**
 * Leitura de FAQ e Linha do Tempo para a visão pública do Colaborador (Faq.html).
 * @return {Object}
 */
function obterDadosFaqETimeline() {
  try {
    var ssChamados = getSpreadsheetChamados();
    var abaFaq = obterAbaFlexivel(ssChamados, CONFIG.ABA_FAQ);
    var abaTimeline = obterAbaFlexivel(ssChamados, CONFIG.ABA_TIMELINE);

    var listaFaq = [];
    var listaTimeline = [];

    if (abaFaq) {
      var dadosFaq = abaFaq.getDataRange().getValues();
      for (var f = 1; f < dadosFaq.length; f++) {
        var status = String(dadosFaq[f][5] || "Ativo").trim().toLowerCase();
        if (status === "ativo") {
          listaFaq.push({
            id: String(dadosFaq[f][0] || ""),
            categoria: String(dadosFaq[f][1] || "Geral").trim(),
            pergunta: String(dadosFaq[f][2] || "").trim(),
            resposta: converterLinksParaHtml(dadosFaq[f][3]),
            ordem: parseInt(dadosFaq[f][4] || 99, 10),
            status: "Ativo"
          });
        }
      }
      listaFaq.sort(function(a, b) { return a.ordem - b.ordem; });
    }

    if (abaTimeline) {
      var dadosTime = abaTimeline.getDataRange().getValues();
      for (var t = 1; t < dadosTime.length; t++) {
        var exibir = String(dadosTime[t][5] || "Sim").trim().toLowerCase();
        if (exibir === "sim" || exibir === "s") {
          listaTimeline.push({
            dataAviso: dadosTime[t][0] ? Utilities.formatDate(new Date(dadosTime[t][0]), "GMT-03:00", "dd/MM/yyyy HH:mm") : "",
            sistemaModulo: String(dadosTime[t][1] || "Sistema GP").trim(),
            tituloAviso: String(dadosTime[t][2] || "").trim(),
            mensagem: converterLinksParaHtml(dadosTime[t][3]),
            statusSistema: String(dadosTime[t][4] || "✅ Estável / Normalizado").trim(),
            exibir: "Sim"
          });
        }
      }
      listaTimeline.reverse();
    }

    return { sucesso: true, faq: listaFaq, timeline: listaTimeline };
  } catch (e) {
    Logger.log("Erro em obterDadosFaqETimeline: " + e.toString());
    return { sucesso: false, faq: [], timeline: [] };
  }
}

/**
 * Leitura COMPLETA de FAQ e Linha do Tempo para o Modal Administrativo do Coordenador.
 * @param {string} [chaveAcesso]
 * @return {Object}
 */
function obterDadosFaqETimelineCompleto(chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Acesso não autorizado." };
  }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaFaq = obterAbaFlexivel(ssChamados, CONFIG.ABA_FAQ);
    var abaTimeline = obterAbaFlexivel(ssChamados, CONFIG.ABA_TIMELINE);

    var listaFaq = [];
    var listaTimeline = [];

    if (abaFaq) {
      var dadosFaq = abaFaq.getDataRange().getValues();
      for (var f = 1; f < dadosFaq.length; f++) {
        if (dadosFaq[f][0] || dadosFaq[f][2]) {
          listaFaq.push({
            id: String(dadosFaq[f][0] || ("FAQ-" + f)),
            categoria: String(dadosFaq[f][1] || "Geral").trim(),
            pergunta: String(dadosFaq[f][2] || "").trim(),
            resposta: String(dadosFaq[f][3] || "").trim(),
            ordem: parseInt(dadosFaq[f][4] || f, 10),
            status: String(dadosFaq[f][5] || "Ativo").trim()
          });
        }
      }
    }

    if (abaTimeline) {
      var dadosTime = abaTimeline.getDataRange().getValues();
      for (var t = 1; t < dadosTime.length; t++) {
        if (dadosTime[t][1] || dadosTime[t][2]) {
          listaTimeline.push({
            id: "TIME-" + t,
            dataAviso: dadosTime[t][0] ? Utilities.formatDate(new Date(dadosTime[t][0]), "GMT-03:00", "dd/MM/yyyy HH:mm") : "",
            sistemaModulo: String(dadosTime[t][1] || "Geral").trim(),
            tituloAviso: String(dadosTime[t][2] || "").trim(),
            mensagem: String(dadosTime[t][3] || "").trim(),
            statusSistema: String(dadosTime[t][4] || "✅ Estável / Normalizado").trim(),
            exibir: String(dadosTime[t][5] || "Sim").trim()
          });
        }
      }
    }

    return { sucesso: true, faq: listaFaq, timeline: listaTimeline };
  } catch (e) {
    Logger.log("Erro em obterDadosFaqETimelineCompleto: " + e.toString());
    return { sucesso: false, mensagem: e.toString() };
  }
}

/**
 * Busca dados do colaborador pelo ID/Matrícula na aba DADOS_ATIVOS.
 * @param {string|number} idInput
 * @return {Object}
 */
function buscarColaboradorPorId(idInput) {
  var idLimpo = String(idInput || "").replace(/\D/g, '').trim();
  if (!idLimpo) {
    return { sucesso: false, mensagem: "Informe um ID / Matrícula válido." };
  }

  try {
    var ssMaster = getSpreadsheetMaster();
    var abaAtivos = obterAbaFlexivel(ssMaster, CONFIG.ABA_ATIVOS);
    
    if (!abaAtivos) {
      return { sucesso: false, mensagem: "Aba DADOS_ATIVOS não foi localizada." };
    }

    var dados = abaAtivos.getDataRange().getValues();
    
    for (var i = 1; i < dados.length; i++) {
      var idBd = String(dados[i][1] || "").replace(/\D/g, '').trim();
      if (idBd === idLimpo) {
        var nomeColab = String(dados[i][2] || "").trim();
        var filialBruta = dados[i][3];
        var filialNormalizada = normalizarFilialId(filialBruta);
        var cargoColab = String(dados[i][8] || dados[i][4] || "Colaborador").trim();

        return {
          sucesso: true,
          colaborador: {
            id: idLimpo,
            nome: nomeColab,
            filialId: filialNormalizada,
            filialNome: "Loja " + filialNormalizada,
            cargo: cargoColab
          }
        };
      }
    }

    return { sucesso: false, mensagem: "Matrícula/ID não localizado na base DADOS_ATIVOS." };
  } catch (e) {
    Logger.log("Erro em buscarColaboradorPorId: " + e.toString());
    return { sucesso: false, mensagem: "Erro ao consultar base de ativos: " + e.toString() };
  }
}

/**
 * Busca informações da filial pelo número.
 * @param {string|number} filialInput
 * @return {Object}
 */
function buscarInfoFilial(filialInput) {
  var filialNorm = normalizarFilialId(filialInput);
  if (!filialNorm) {
    return { sucesso: false, filialId: "", filialNome: "" };
  }
  return { sucesso: true, filialId: filialNorm, filialNome: "Loja " + filialNorm };
}

/**
 * Lista as categorias de chamados ativas.
 * @return {Array<string>}
 */
function obterListaCategorias() {
  try {
    var ssChamados = getSpreadsheetChamados();
    var abaCat = obterAbaFlexivel(ssChamados, CONFIG.ABA_CATEGORIAS);
    
    if (abaCat) {
      var dados = abaCat.getDataRange().getValues();
      var categorias = [];
      for (var i = 1; i < dados.length; i++) {
        var cat = String(dados[i][0] || "").trim();
        if (cat !== "") {
          categorias.push(cat);
        }
      }
      if (categorias.length > 0) return categorias;
    }
  } catch (e) {
    Logger.log("Erro ao carregar categorias dinâmicas: " + e.toString());
  }
  return CONFIG.CATEGORIAS_PADRAO;
}

/**
 * Lista chamados para o painel de atendimento dos Coordenadores.
 * @param {string} [filtroStatus]
 * @param {string} [chaveAcesso]
 * @return {Object}
 */
function listarChamadosPainel(filtroStatus, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return {
      sucesso: false,
      mensagem: "Acesso restrito ao GP.",
      metricas: { total: 0, abertos: 0, emAnalise: 0, concluidos: 0 },
      chamados: [],
      kpiCategorias: []
    };
  }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    
    if (!abaBase) {
      return {
        sucesso: true,
        chamados: [],
        metricas: { total: 0, abertos: 0, emAnalise: 0, concluidos: 0 },
        kpiCategorias: []
      };
    }

    var dados = abaBase.getDataRange().getValues();
    var chamados = [];
    var statusFiltro = String(filtroStatus || "Todos").trim().toLowerCase();

    var mapaCargos = obterMapaCargosAtivos();

    var totalAbertos = 0;
    var totalEmAnalise = 0;
    var totalConcluidos = 0;
    var contagemCategorias = {};

    if (dados.length <= 1) {
      return {
        sucesso: true,
        chamados: [],
        metricas: { total: 0, abertos: 0, emAnalise: 0, concluidos: 0 },
        kpiCategorias: []
      };
    }

    var dataAtual = new Date();

    for (var i = 1; i < dados.length; i++) {
      var prot = String(dados[i][0] || "").trim();
      if (!prot || prot.indexOf("GP-") !== 0) continue; // Filtra e ignora linhas invalidadas

      var st = String(dados[i][9] || "Novo / Aberto").trim();
      var cat = String(dados[i][6] || "Outros").trim();
      var resp = String(dados[i][11] || "Não atribuído").trim();
      var ultInteracao = dados[i][10] ? new Date(dados[i][10]) : null;

      var emAtendimentoPorOutro = false;
      if (st.indexOf("Análise") !== -1 && resp !== "Não atribuído" && ultInteracao) {
        var diffMinutos = Math.floor((dataAtual - ultInteracao) / (1000 * 60));
        if (diffMinutos < 15) {
          emAtendimentoPorOutro = true;
        }
      }

      var stLower = st.toLowerCase();

      if (stLower.indexOf("novo") !== -1 || stLower.indexOf("aberto") !== -1) totalAbertos++;
      else if (stLower.indexOf("análise") !== -1 || stLower.indexOf("analise") !== -1) totalEmAnalise++;
      else if (stLower.indexOf("respondido") !== -1 || stLower.indexOf("concluído") !== -1 || stLower.indexOf("concluido") !== -1) totalConcluidos++;

      contagemCategorias[cat] = (contagemCategorias[cat] || 0) + 1;

      var passaFiltro = false;
      if (statusFiltro === "todos") {
        passaFiltro = true;
      } else if (statusFiltro.indexOf("novo") !== -1 && (stLower.indexOf("novo") !== -1 || stLower.indexOf("aberto") !== -1)) {
        passaFiltro = true;
      } else if (statusFiltro.indexOf("análise") !== -1 && (stLower.indexOf("análise") !== -1 || stLower.indexOf("analise") !== -1)) {
        passaFiltro = true;
      } else if (statusFiltro.indexOf("concluído") !== -1 && (stLower.indexOf("respondido") !== -1 || stLower.indexOf("concluído") !== -1 || stLower.indexOf("concluido") !== -1)) {
        passaFiltro = true;
      }

      if (passaFiltro) {
        var idSol = String(dados[i][2] || "").replace(/\D/g, '').trim();
        var cargoSol = mapaCargos[idSol] || "Colaborador";
        var filialNorm = normalizarFilialId(dados[i][4]);

        chamados.push({
          protocolo: prot,
          dataAbertura: dados[i][1] ? Utilities.formatDate(new Date(dados[i][1]), "GMT-03:00", "dd/MM/yyyy HH:mm") : "",
          solicitanteEmail: "Não informado",
          solicitanteId: idSol,
          solicitanteNome: String(dados[i][3] || "").trim(),
          solicitanteCargo: cargoSol,
          filialId: filialNorm,
          filialNome: String(dados[i][5] || ("Loja " + filialNorm)).trim(),
          categoria: cat,
          assunto: String(dados[i][7] || "").trim(),
          descricao: converterLinksParaHtml(dados[i][8]),
          status: st,
          atribuidoA: resp,
          emAtendimento: emAtendimentoPorOutro,
          ultimaAtualizacao: dados[i][10] ? Utilities.formatDate(new Date(dados[i][10]), "GMT-03:00", "dd/MM/yyyy HH:mm") : ""
        });
      }
    }

    chamados.reverse();

    var arrayKpiCategorias = [];
    var totalGeralChamados = totalAbertos + totalEmAnalise + totalConcluidos;
    for (var c in contagemCategorias) {
      var qtd = contagemCategorias[c];
      var pct = totalGeralChamados > 0 ? Math.round((qtd / totalGeralChamados) * 100) : 0;
      arrayKpiCategorias.push({ categoria: c, quantidade: qtd, percentual: pct });
    }
    arrayKpiCategorias.sort(function(a, b) { return b.quantidade - a.quantidade; });

    return {
      sucesso: true,
      chamados: chamados,
      usuarioLogado: permissao.email,
      metricas: {
        total: totalGeralChamados,
        abertos: totalAbertos,
        emAnalise: totalEmAnalise,
        concluidos: totalConcluidos
      },
      kpiCategorias: arrayKpiCategorias
    };
  } catch (e) {
    Logger.log("Erro em listarChamadosPainel: " + e.toString());
    return {
      sucesso: false,
      mensagem: "Erro ao listar chamados: " + e.toString(),
      metricas: { total: 0, abertos: 0, emAnalise: 0, concluidos: 0 },
      chamados: [],
      kpiCategorias: []
    };
  }
}

/**
 * Consulta a lista de dúvidas enviadas por um colaborador.
 * @param {string|number} idInput
 * @return {Object}
 */
function listarChamadosPorColaboradorId(idInput) {
  var idLimpo = String(idInput || "").replace(/\D/g, '').trim();
  if (!idLimpo) {
    return { sucesso: false, mensagem: "Informe uma Matrícula / ID válida." };
  }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);

    if (!abaBase) {
      return { sucesso: true, chamados: [] };
    }

    var dadosBase = abaBase.getDataRange().getValues();
    var dadosHist = abaHist ? abaHist.getDataRange().getValues() : [];
    
    var listaChamados = [];

    for (var i = 1; i < dadosBase.length; i++) {
      var idBd = String(dadosBase[i][2] || "").replace(/\D/g, '').trim();
      
      if (idBd === idLimpo) {
        var prot = String(dadosBase[i][0] || "").trim();
        var historicoMensagens = [];

        for (var h = 1; h < dadosHist.length; h++) {
          if (String(dadosHist[h][0] || "").trim() === prot) {
            historicoMensagens.push({
              dataHora: dadosHist[h][1] ? Utilities.formatDate(new Date(dadosHist[h][1]), "GMT-03:00", "dd/MM/yyyy HH:mm:ss") : "",
              autorNome: dadosHist[h][2],
              autorPapel: dadosHist[h][3],
              mensagem: converterLinksParaHtml(dadosHist[h][4])
            });
          }
        }

        var filialNorm = normalizarFilialId(dadosBase[i][4]);

        listaChamados.push({
          protocolo: prot,
          dataAbertura: dadosBase[i][1] ? Utilities.formatDate(new Date(dadosBase[i][1]), "GMT-03:00", "dd/MM/yyyy HH:mm") : "",
          solicitanteNome: dadosBase[i][3],
          filialId: filialNorm,
          filialNome: dadosBase[i][5] || ("Loja " + filialNorm),
          categoria: dadosBase[i][6],
          assunto: dadosBase[i][7],
          descricao: converterLinksParaHtml(dadosBase[i][8]),
          status: dadosBase[i][9],
          historico: historicoMensagens
        });
      }
    }

    listaChamados.reverse();

    return {
      sucesso: true,
      colaboradorId: idLimpo,
      chamados: listaChamados
    };
  } catch (e) {
    Logger.log("Erro em listarChamadosPorColaboradorId: " + e.toString());
    return { sucesso: false, mensagem: "Erro ao buscar chamados: " + e.toString() };
  }
}

/**
 * Consulta histórico por protocolo.
 * @param {string} protocolo
 * @param {string} [chaveAcesso]
 * @return {Object}
 */
function consultarHistoricoChamado(protocolo, chaveAcesso) {
  var permissao = validarAcessoCoordenador(chaveAcesso);
  if (!permissao.autorizado) {
    return { sucesso: false, mensagem: "Acesso não autorizado." };
  }

  var protLimpo = String(protocolo || "").trim();
  if (!protLimpo) return { sucesso: false, mensagem: "Protocolo inválido." };

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);
    var historico = [];

    if (abaHist) {
      var dados = abaHist.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        if (String(dados[i][0] || "").trim() === protLimpo) {
          historico.push({
            dataHora: dados[i][1] ? Utilities.formatDate(new Date(dados[i][1]), "GMT-03:00", "dd/MM/yyyy HH:mm:ss") : "",
            autorNome: dados[i][2],
            autorPapel: dados[i][3],
            mensagem: converterLinksParaHtml(dados[i][4])
          });
        }
      }
    }

    return { sucesso: true, protocolo: protLimpo, historico: historico };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao consultar histórico: " + e.toString() };
  }
}

/**
 * Registra réplica do colaborador.
 * @param {string} protocolo
 * @param {string} matricula
 * @param {string} texto
 * @return {Object}
 */
function enviarReplicaColaborador(protocolo, matricula, texto) {
  var protLimpo = String(protocolo || "").trim();
  var matLimpa = String(matricula || "").replace(/\D/g, '').trim();
  var textoLimpo = String(texto || "").trim();

  if (!protLimpo || !matLimpa || !textoLimpo) {
    return { sucesso: false, mensagem: "Preencha todos os campos para enviar a réplica." };
  }

  try {
    var ssChamados = getSpreadsheetChamados();
    var abaBase = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_BASE);
    var abaHist = obterAbaFlexivel(ssChamados, CONFIG.ABA_CHAMADOS_HISTORICO);

    if (!abaBase || !abaHist) {
      return { sucesso: false, mensagem: "Base de chamados indisponível." };
    }

    var dadosBase = abaBase.getDataRange().getValues();
    var linhaEncontrada = -1;
    var nomeSolicitante = "Colaborador";

    for (var i = 1; i < dadosBase.length; i++) {
      var protBd = String(dadosBase[i][0] || "").trim();
      var idBd = String(dadosBase[i][2] || "").replace(/\D/g, '').trim();

      if (protBd === protLimpo && idBd === matLimpa) {
        linhaEncontrada = i + 1;
        nomeSolicitante = String(dadosBase[i][3] || "Colaborador").trim();
        break;
      }
    }

    if (linhaEncontrada === -1) {
      return { sucesso: false, mensagem: "Dúvida não encontrada para esta matrícula." };
    }

    var dataAtual = new Date();

    abaBase.getRange(linhaEncontrada, 10).setValue("Em Análise");
    abaBase.getRange(linhaEncontrada, 11).setValue(dataAtual);

    abaHist.appendRow([
      protLimpo,
      dataAtual,
      nomeSolicitante,
      "Solicitante",
      textoLimpo
    ]);

    return { sucesso: true, mensagem: "Sua mensagem complementar foi enviada com sucesso!" };
  } catch (e) {
    Logger.log("Erro em enviarReplicaColaborador: " + e.toString());
    return { sucesso: false, mensagem: "Erro ao salvar mensagem: " + e.toString() };
  }
}

/**
 * Gera mapa de cargos ativos.
 * @return {Object}
 */
function obterMapaCargosAtivos() {
  var mapa = {};
  try {
    var ssMaster = getSpreadsheetMaster();
    var abaAtivos = obterAbaFlexivel(ssMaster, CONFIG.ABA_ATIVOS);
    if (abaAtivos) {
      var dados = abaAtivos.getDataRange().getValues();
      for (var i = 1; i < dados.length; i++) {
        var id = String(dados[i][1] || "").replace(/\D/g, '').trim();
        var cargo = String(dados[i][8] || dados[i][4] || "").trim();
        if (id && cargo) {
          mapa[id] = cargo;
        }
      }
    }
  } catch (e) {
    Logger.log("Erro ao carregar mapa de cargos: " + e.toString());
  }
  return mapa;
}
