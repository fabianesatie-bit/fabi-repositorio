// =============================================================================
// CONFIGURAÇÕES GLOBAIS DO BANCO DE DADOS GP - VERSÃO BLINDADA E OTIMIZADA
// =============================================================================

const SPREADSHEET_ID = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; // ID Planilha GP 360 (DB_MASTER)
const EVIDENCIAS_FOLDER_ID = '1v28G-ZDd6yQpjTUvBNcODlpxkM5AeaQZ'; // ID Pasta do Drive para Evidências
const SPREADSHEET_DASH = '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc'; // Base de Indicadores Geral (DB_DASH)
const SPREADSHEET_LOG_ID = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY'; // ID Planilha Auditoria Exclusiva

// Lista de administradores com acesso irrestrito
const SUPER_ADMINS_EMAILS = [
  "fabiane.satie@magazineluiza.com.br",
  "gplojas@magazineluiza.com.br",
  "tarcisio.maniglia@magazineluiza.com.br"
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal GP 360')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// MOTOR DE AUDITORIA CENTRALIZADO
function registrarAuditoria(evento, detalhe) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const ssLog = SpreadsheetApp.openById(SPREADSHEET_LOG_ID);
    const sheetLog = ssLog.getSheetByName('AUDITORIA') || ssLog.getSheets()[0];
    sheetLog.appendRow([new Date(), emailLogado, evento, detalhe]);
  } catch(e) { /* Falha silenciosa para estabilidade da UX */ }
}

// =============================================================================
// MOTOR DE SEGURANÇA: CONTROLE DE ACESSO E CONTEXTO ATIVO (SERVER-SIDE)
// =============================================================================

function obterControleAcesso(email) {
  if (!email) return { autorizado: false, erro: "Sessão não identificada." };

  const emailNorm = email.toLowerCase().trim();
  const isSuper = SUPER_ADMINS_EMAILS.includes(emailNorm);

  if (isSuper) {
    return {
      autorizado: true,
      email: emailNorm,
      nome: emailNorm === "fabiane.satie@magazineluiza.com.br" ? "FABIANE SATIE" : "SUPER ADMIN GP",
      cargo: "Administrador",
      isSuperAdmin: true,
      isAdmin: true,
      isConfigAdmin: true,
      isGerenteGP: true,
      regionais: ["TODAS"],
      diretoriasAtendidas: ["TODAS"]
    };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetUsuarios = ss.getSheetByName('DADOS_USUARIOS');
    if (!sheetUsuarios) return { autorizado: false, erro: "Base de usuários inacessível." };
    const dadosUsuarios = sheetUsuarios.getDataRange().getValues();
    for (let i = 1; i < dadosUsuarios.length; i++) {
      const emailLinha = dadosUsuarios[i][0] ? String(dadosUsuarios[i][0]).toLowerCase().trim() : "";
      if (emailLinha === emailNorm) {
        const status = dadosUsuarios[i][8] ? String(dadosUsuarios[i][8]).toLowerCase().trim() : "ativo";
        if (status === "inativo" || status === "desligado") {
          return { autorizado: false, erro: "Sessão inativa ou revogada." };
        }
        const cargo = dadosUsuarios[i][2] || "Cargo não definido";
        const nivelAcesso = dadosUsuarios[i][5] ? String(dadosUsuarios[i][5]).trim() : "Sem acesso";
        if (nivelAcesso.toLowerCase() === "sem acesso") {
          return { autorizado: false, erro: "Usuário sem permissão de acesso ao Portal." };
        }
        const adminRoles = ["GerenteGP", "Administrador", "DiretorRH"];
        const isAdmin = adminRoles.includes(cargo) || adminRoles.includes(nivelAcesso);
        const isGerenteGP = (cargo === "GerenteGP" || nivelAcesso === "GerenteGP" || isAdmin);

        return {
          autorizado: true,
          email: emailNorm,
          nome: dadosUsuarios[i][1] || "Nome não cadastrado",
          cargo: cargo,
          isSuperAdmin: (cargo === "Administrador" || nivelAcesso === "Administrador"),
          isAdmin: isAdmin,
          isConfigAdmin: isAdmin,
          isGerenteGP: isGerenteGP,
          regionais: dadosUsuarios[i][4] ? String(dadosUsuarios[i][4]).split(',').map(r => r.trim()) : [],
          diretoriasAtendidas: dadosUsuarios[i][3] ? String(dadosUsuarios[i][3]).split(',').map(d => d.trim()) : []
        };
      }
    }
  } catch (e) {
    return { autorizado: false, erro: "Falha de comunicação com o servidor de acessos." };
  }
  return { autorizado: false, erro: "Usuário não localizado na base corporativa." };
}

// =============================================================================
// HELPERS E MÉTODOS DE EXTRAÇÃO SEGURA
// =============================================================================

function extrairIdDrive(linkOuId) {
  if (!linkOuId) return "";
  const match = linkOuId.match(/[-\w]{25,}(?!.*[-\w]{25,})/);
  return match ? match[0] : linkOuId;
}

function formatarDataSegura(dataValor) {
  if (!dataValor) return "";
  try { 
    if (dataValor instanceof Date) {
        return Utilities.formatDate(dataValor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) return ds;

    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) {
        return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); 
    }
    return ds; 
  } catch(e) { return String(dataValor).split(' ')[0]; }
}

function obterDataRawSegura(dataValor) {
  if (!dataValor) return 0;
  try { 
    if (dataValor instanceof Date) return dataValor.getTime();

    let str = String(dataValor).trim();
    let ds = str.split(' ')[0];
    if (ds.includes('/')) {
        let p = ds.split('/');
        if(p.length === 3) return new Date(p[2], p[1]-1, p[0]).getTime();
    } else if (ds.includes('-')) {
        let p = ds.split('-');
        if(p.length === 3) return new Date(p[0], p[1]-1, p[2]).getTime();
    }
    const dt = new Date(dataValor); 
    if (!isNaN(dt.getTime())) return dt.getTime(); 
    return 0; 
  } catch(e) { return 0; }
}

// =============================================================================
// CARREGAMENTO DE DADOS E RELACIONAMENTOS (PORTAL HOME & GAMIFICAÇÃO)
// =============================================================================

function obterDadosIniciais() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      return { bloqueado: true, mensagem: controle.erro };
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let usuario = { 
      email: controle.email, 
      nome: controle.nome, 
      cargo: controle.cargo, 
      infoLinha2: controle.isSuperAdmin ? "Visão Nacional" : controle.regionais.join(', '), 
      regionais: controle.regionais, 
      diretoriasAtendidas: controle.diretoriasAtendidas || [],
      isAdmin: controle.isAdmin, 
      isSuperAdmin: controle.isSuperAdmin, 
      isConfigAdmin: controle.isConfigAdmin, 
      isGerenteGP: controle.isGerenteGP,
      foto: "" 
    };
    let mapaUsuarios = {}; 
    const sheetUsuarios = ss.getSheetByName('DADOS_USUARIOS');
    let listaDiretoriasUnicas = new Set();
    if (sheetUsuarios) {
      const dadosUsuarios = sheetUsuarios.getDataRange().getValues();
      for (let i = 1; i < dadosUsuarios.length; i++) {
        const emailLinha = dadosUsuarios[i][0] ? String(dadosUsuarios[i][0]).toLowerCase().trim() : "";
        if (emailLinha) {
           let dirRaw = dadosUsuarios[i][3] ? String(dadosUsuarios[i][3]).trim() : "";
           if (dirRaw) {
              dirRaw.split(',').forEach(d => listaDiretoriasUnicas.add(d.trim()));
           }
           let fotoRaw = (dadosUsuarios[i].length > 6 && dadosUsuarios[i][6]) ? String(dadosUsuarios[i][6]).trim() : "";
           let fotoId = extrairIdDrive(fotoRaw);
           let fotoPronta = fotoId ? (fotoId.startsWith("http") ? fotoId : "https://lh3.googleusercontent.com/d/" + fotoId) : "";

           mapaUsuarios[emailLinha] = { 
               nome: dadosUsuarios[i][1] || 'GP', 
               foto: fotoPronta,
               regionais: dadosUsuarios[i][4] ? String(dadosUsuarios[i][4]).split(',').map(r => r.trim()) : []
           };
        }
        if (emailLinha === emailLogado) {
          usuario.foto = mapaUsuarios[emailLinha].foto;
        }
      }
    }
    registrarAuditoria("Login", "Acesso ao Portal Concluído");

    let placarMoedasTotal = {};
    let placarMoedasMes = {};
    let gastoTotalAcumulado = 0;
    let setVisitasFisicas = new Set();

    const dHoje = new Date();
    const mesAtual = dHoje.getMonth();
    const anoAtual = dHoje.getFullYear();
    const meusLancamentos = []; 
    const sheetLancamentos = ss.getSheetByName('DADOS_LANCAMENTOS'); 

    if (sheetLancamentos && sheetLancamentos.getLastRow() > 1) {
      const dadosLancamentos = sheetLancamentos.getDataRange().getValues();
      for (let i = 1; i < dadosLancamentos.length; i++) {
        const emailAutor = dadosLancamentos[i][2] ? String(dadosLancamentos[i][2]).toLowerCase().trim() : "";
        if (emailAutor) {
           if (!placarMoedasTotal[emailAutor]) placarMoedasTotal[emailAutor] = 0;
           if (!placarMoedasMes[emailAutor]) placarMoedasMes[emailAutor] = 0;

           let valorMoedaRaw = String(dadosLancamentos[i][9] || "0").replace(',', '.').trim();
           let moedasParse = parseFloat(valorMoedaRaw);
           if (!isNaN(moedasParse)) {
               placarMoedasTotal[emailAutor] += moedasParse;

               let dtAcaoRaw = dadosLancamentos[i][12] || dadosLancamentos[i][1];
               if (dtAcaoRaw) {
                 let dtObj = new Date(dtAcaoRaw);
                 if (!isNaN(dtObj.getTime()) && dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual) {
                   placarMoedasMes[emailAutor] += moedasParse;
                 }
               }
           }
        }
        let visivelNoHistorico = false;
        if (usuario.isSuperAdmin) {
          visivelNoHistorico = true;
        } else if (usuario.isAdmin) {
          let regAutor = mapaUsuarios[emailAutor] ? mapaUsuarios[emailAutor].regionais : [];
          let overlap = regAutor.some(r => usuario.regionais.includes(r));
          if (overlap || emailAutor === emailLogado) visivelNoHistorico = true;
        } else {
          if (emailAutor === emailLogado) visivelNoHistorico = true;
        }
        if (visivelNoHistorico) {
          let kmCusto = Number(dadosLancamentos[i][16]) || 0;
          let alimentacao = Number(dadosLancamentos[i][22]) || 0;
          let hospedagem = Number(dadosLancamentos[i][23]) || 0;
          let aereo = Number(dadosLancamentos[i][24]) || 0;
          let pedagio = Number(dadosLancamentos[i][25]) || 0;
          let estacionamento = Number(dadosLancamentos[i][26]) || 0;
          let gastoSomaReal = kmCusto + alimentacao + hospedagem + aereo + pedagio + estacionamento;

          let dtViagemReal = obterDataRawSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
          if (emailAutor === emailLogado) {
              let dtObj = new Date(dtViagemReal);
              if (!isNaN(dtObj.getTime())) {
                  if (dtObj.getMonth() === mesAtual && dtObj.getFullYear() === anoAtual) {
                      gastoTotalAcumulado += gastoSomaReal;
                  }
              }
              let rot = String(dadosLancamentos[i][17] || "").trim();
              if (rot.includes("Visita in loco")) {
                  let dtVFormat = formatarDataSegura(dadosLancamentos[i][12] || dadosLancamentos[i][1]);
                  let dst = String(dadosLancamentos[i][3]).trim();
                  setVisitasFisicas.add(emailAutor + "|" + dtVFormat + "|" + dst);
              }
          }
          meusLancamentos.push({
            id: dadosLancamentos[i][0], 
            dataRegistro: formatarDataSegura(dadosLancamentos[i][1]), 
            autorNome: mapaUsuarios[emailAutor] ? mapaUsuarios[emailAutor].nome : emailAutor,
            isOwner: (emailAutor === emailLogado), 
            destino: dadosLancamentos[i][3], 
            motivo: dadosLancamentos[i][4], 
            gastoTotal: gastoSomaReal, 
            dataViagem: formatarDataSegura(dadosLancamentos[i][12]), 
            dataViagemRaw: dtViagemReal, 
            kmValor: Number(dadosLancamentos[i][14]) || 0, 
            kmQtd: Number(dadosLancamentos[i][15]) || 0, 
            kmCusto: kmCusto, 
            alimentacao: alimentacao,
            hospedagem: hospedagem,
            aereo: aereo,
            pedagio: pedagio,
            estacionamento: estacionamento,
            observacoes: String(dadosLancamentos[i][10] || "").trim(), 
            roteiro: String(dadosLancamentos[i][17] || "").trim(), 
            subTema: String(dadosLancamentos[i][18] || "").trim(), 
            pessoasImpactadas: Number(dadosLancamentos[i][19]) || 0, 
            tempoGasto: Number(dadosLancamentos[i][20]) || 0,
            linkEvidencia: String(dadosLancamentos[i][11] || "").trim()
          });
        }
      }
    }

    meusLancamentos.reverse(); 
    let historicoLimitado = meusLancamentos.slice(0, 500);

    let premiosDicionario = getDicionarioPremios();
    let userMoedasTotal = placarMoedasTotal[emailLogado] || 0;
    let userMoedasMes = placarMoedasMes[emailLogado] || 0;

    let badgesConquistadas = premiosDicionario.filter(p => userMoedasTotal >= p.meta && p.meta > 0);

    // Regra unificada de Fase baseada no saldo exibido
    let faseMontanha = 1;
    if (userMoedasTotal >= 120) faseMontanha = 4;
    else if (userMoedasTotal >= 80) faseMontanha = 3;
    else if (userMoedasTotal >= 40) faseMontanha = 2;

    let ranking = Object.keys(placarMoedasTotal).map(email => {
        let nomeExibicao = 'GP';
        let fotoExibicao = '';
        if (mapaUsuarios[email]) {
            nomeExibicao = mapaUsuarios[email].nome;
            fotoExibicao = mapaUsuarios[email].foto;
        } else {
            let partesEmail = email.split('@')[0].split('.');
            nomeExibicao = partesEmail.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
        }
        let totalU = placarMoedasTotal[email] || 0;
        let badgesU = premiosDicionario.filter(p => totalU >= p.meta && p.meta > 0).map(p => p.icone);

        // Regra unificada e coerente para o ranking
        let faseU = 1;
        if (totalU >= 120) faseU = 4;
        else if (totalU >= 80) faseU = 3;
        else if (totalU >= 40) faseU = 2;

        return { 
          email: email, 
          nome: nomeExibicao, 
          foto: fotoExibicao, 
          moedas: totalU,
          moedasMes: placarMoedasMes[email] || 0,
          fase: faseU,
          badges: badgesU
        };
    });
    ranking = ranking.filter(user => user.moedas > 0 && !isNaN(user.moedas));
    ranking.sort((a, b) => b.moedas - a.moedas);

    const lojasFiltradasMap = new Map();
    const sheetLojas = ss.getSheetByName('DADOS_LOJAS');
    if (sheetLojas && sheetLojas.getLastRow() > 1) {
      const dadosLojas = sheetLojas.getDataRange().getValues();
      for (let i = 1; i < dadosLojas.length; i++) {
        const regionalLoja = dadosLojas[i][2] ? String(dadosLojas[i][2]).trim() : "";
        if (usuario.isSuperAdmin || usuario.regionais.includes(regionalLoja)) {
          let idLoja = parseInt(dadosLojas[i][0], 10);
          if (!isNaN(idLoja)) {
              if (idLoja > 3000) idLoja = idLoja - 3000;
              if (!lojasFiltradasMap.has(idLoja)) {
                  lojasFiltradasMap.set(idLoja, { id: idLoja, nome: dadosLojas[i][1], regional: regionalLoja });
              }
          }
        }
      }
    }
    const lojasFiltradas = Array.from(lojasFiltradasMap.values());
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    const avisos = [];
    if (sheetAvisos && sheetAvisos.getLastRow() > 1) {
       const dadosAvisos = sheetAvisos.getDataRange().getValues();
       for(let i = 1; i < dadosAvisos.length; i++) {
          avisos.push({ data: formatarDataSegura(dadosAvisos[i][0]), autor: dadosAvisos[i][1], mensagem: dadosAvisos[i][2] });
       }
       avisos.reverse(); 
    }
    let naturezas = carregarNaturezasSeguras(ss);
    let diretoriasAr = Array.from(listaDiretoriasUnicas).filter(Boolean).sort();

    return { 
      bloqueado: false, usuario: usuario, lojas: lojasFiltradas, qtdLojasCarteira: lojasFiltradas.length,
      qtdVisitas: setVisitasFisicas.size,
      moedas: userMoedasTotal,
      moedasMes: userMoedasMes,
      faseMontanha: faseMontanha,
      badges: badgesConquistadas,
      gastos: gastoTotalAcumulado, historico: historicoLimitado, 
      avisos: avisos, ranking: ranking.slice(0, 5), naturezas: naturezas, diretorias: diretoriasAr
    };
  } catch (e) { return { erro: e.message }; }
}

function carregarNaturezasSeguras(ss) {
  let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
  if (!sheetConfig) return [];
  const valores = sheetConfig.getDataRange().getValues();
  // Removido "Multiplicação de Conhecimento" e "Tech Talk / Compartilhamento" da lista fixa conforme solicitado
  const fixas = ["Reunião regional", "Treinamentos", "Celebrações/Ritão/Reconhecimento", "Recrutamento e seleção - 1º liderança", "Recrutamento e seleção - Processo Externo", "Reunião Conselho E Conselho consultivo", "NPS - Lojas", "Receita de Mercadoria", "GMD - Operações de Loja", "Relatorios/Feedback/Acompanhamento", "Atendimento Social"];
  let list = [];
  for (let i = 1; i < valores.length; i++) {
    let item = valores[i][0] ? String(valores[i][0]).trim() : "";
    if (item && item !== 'Naturezas_Atividades' && !fixas.includes(item)) list.push(item);
  }
  return list;
}

// =============================================================================
// GALERIA DE EVIDÊNCIAS E BYPASS CORS BASE64 (HIERARQUIA BLINDADA)
// =============================================================================

function getEvidenciasUsuario() {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLancamentos = ss.getSheetByName('DADOS_LANCAMENTOS');
    const sheetUsu = ss.getSheetByName('DADOS_USUARIOS');
    if (!sheetLancamentos || !sheetUsu) return [];

    let mapUsu = {};
    const dUsu = sheetUsu.getDataRange().getValues();
    for(let i=1; i<dUsu.length; i++){
        let em = String(dUsu[i][0]).toLowerCase().trim();
        let dirArr = dUsu[i][3] ? String(dUsu[i][3]).toLowerCase().split(',').map(s=>s.trim()) : [];
        let regArr = dUsu[i][4] ? String(dUsu[i][4]).toLowerCase().split(',').map(s=>s.trim()) : [];
        mapUsu[em] = { dir: dirArr, reg: regArr };
    }

    const dados = sheetLancamentos.getDataRange().getValues();
    let evidencias = [];

    let myDir = controle.diretoriasAtendidas ? controle.diretoriasAtendidas.map(d=>d.toLowerCase().trim()) : [];
    let myReg = controle.regionais ? controle.regionais.map(r=>r.toLowerCase().trim()) : [];

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const emailAutor = String(linha[2] || "").toLowerCase().trim();
      const linkEvidencia = String(linha[11] || "").trim();
      if (!linkEvidencia) continue;

      let hasAccess = false;
      if (controle.isSuperAdmin) {
          hasAccess = true; 
      } else if (controle.isGerenteGP) {
          if (emailAutor === emailLogado) {
              hasAccess = true;
          } else if (mapUsu[emailAutor]) {
              let overlapDir = mapUsu[emailAutor].dir.some(d => myDir.includes(d));
              let overlapReg = mapUsu[emailAutor].reg.some(r => myReg.includes(r));
              if (overlapDir || overlapReg) hasAccess = true;
          }
      } else {
          if (emailAutor === emailLogado) hasAccess = true; 
      }

      if (hasAccess) {
        const match = linkEvidencia.match(/id=([a-zA-Z0-9_-]+)/) || linkEvidencia.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          const fileId = match[1];
          try {
            let thumbUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w400`;
            evidencias.push({
              id: fileId,
              nome: "Evidência Visual", // Nome genérico injetado para evitar timeout do DriveApp
              motivo: String(linha[4] || "Ação de Governança"),
              destino: String(linha[3] || "Filial"),
              data: formatarDataSegura(linha[12] || linha[1]),
              thumbUrl: thumbUrl,
              downloadUrl: `https://drive.google.com/uc?export=download&id=${fileId}`
            });
          } catch (e) {
            continue;
          }
        }
      }
    }
    return evidencias.reverse().slice(0, 150); // Limite injetado para performance do DOM
  } catch (e) {
    return [];
  }
}

function getImagemBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);

    return {
      sucesso: true,
      mimeType: blob.getContentType(),
      base64: base64
    };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

// =============================================================================
// GESTÃO DO DICIONÁRIO DE PRÊMIOS (GERENTEGP CRUD)
// =============================================================================

function getDicionarioPremios() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');

    if (!abaPremios) {
      abaPremios = ss.insertSheet('DICIONARIO_PREMIOS');
      abaPremios.appendRow(['ID_Premio', 'Nome_Premio', 'Custo_ou_Meta', 'Icone']);
      abaPremios.getRange(1, 1, 1, 4).setFontWeight('bold');
    }

    const dados = abaPremios.getDataRange().getValues();
    let lista = [];

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0]) {
        lista.push({
          id: String(dados[i][0]),
          nome: String(dados[i][1] || ""),
          meta: Number(dados[i][2]) || 0,
          icone: String(dados[i][3] || "🏆")
        });
      }
    }
    return lista;
  } catch (e) {
    return [];
  }
}

function salvarPremio(premio) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.isGerenteGP && !controle.isSuperAdmin) {
      return { sucesso: false, erro: "Acesso negado: Perfil sem permissão de GerenteGP." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
    if (!abaPremios) {
      abaPremios = ss.insertSheet('DICIONARIO_PREMIOS');
      abaPremios.appendRow(['ID_Premio', 'Nome_Premio', 'Custo_ou_Meta', 'Icone']);
    }

    const novoId = premio.id || 'PRM-' + new Date().getTime();
    const dados = abaPremios.getDataRange().getValues();
    let linhaExistente = -1;

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(novoId)) {
        linhaExistente = i + 1;
        break;
      }
    }

    if (linhaExistente > 0) {
      abaPremios.getRange(linhaExistente, 2).setValue(premio.nome);
      abaPremios.getRange(linhaExistente, 3).setValue(premio.meta);
      abaPremios.getRange(linhaExistente, 4).setValue(premio.icone);
    } else {
      abaPremios.appendRow([novoId, premio.nome, premio.meta, premio.icone]);
    }

    return { sucesso: true, premios: getDicionarioPremios() };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

function deletarPremio(idPremio) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.isGerenteGP && !controle.isSuperAdmin) {
      return { sucesso: false, erro: "Acesso negado." };
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const abaPremios = ss.getSheetByName('DICIONARIO_PREMIOS');
    if (!abaPremios) return { sucesso: false, erro: "Aba não encontrada." };

    const dados = abaPremios.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idPremio)) {
        abaPremios.deleteRow(i + 1);
        return { sucesso: true, premios: getDicionarioPremios() };
      }
    }

    return { sucesso: false, erro: "Prêmio não encontrado." };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

// =============================================================================
// VERIFICAÇÃO E APLICAÇÃO DE RLS NAS CONSULTAS DE FILIAIS
// =============================================================================

function validarAcessoFilial(email, filialId) {
  const controle = obterControleAcesso(email);
  if (!controle.autorizado) return { autorizado: false };
  if (controle.isSuperAdmin) return { autorizado: true, controle: controle };
  try {
    let targetFilial = parseInt(filialId, 10);
    if (isNaN(targetFilial)) return { autorizado: false };
    if (targetFilial > 3000) targetFilial -= 3000;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetLojas = ss.getSheetByName('DADOS_LOJAS');
    if (!sheetLojas) return { autorizado: false };
    const dadosLojas = sheetLojas.getDataRange().getValues();
    for (let i = 1; i < dadosLojas.length; i++) {
      let idLoja = parseInt(dadosLojas[i][0], 10);
      if (idLoja > 3000) idLoja -= 3000;

      if (idLoja === targetFilial) {
        const regionalLoja = dadosLojas[i][2] ? String(dadosLojas[i][2]).trim() : "";
        const autorizado = controle.regionais.includes(regionalLoja);
        return { autorizado: autorizado, controle: controle };
      }
    }
  } catch (e) {}
  return { autorizado: false };
}

// =============================================================================
// CONSULTA SEGURA DE INDICADORES (DASHBOARD LOJA)
// =============================================================================

function buscarIndicadoresLoja(filialId) {
  try {
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const verificacao = validarAcessoFilial(emailLogado, filialId);
    if (!verificacao.autorizado) {
      return { erro: "ACESSO NEGADO: Você não tem permissão para ler dados desta filial." };
    }
    let targetFilial = parseInt(filialId, 10);
    if (targetFilial > 3000) targetFilial -= 3000;
    const cache = CacheService.getScriptCache();
    const cacheKey = 'IND_LOJA_' + targetFilial;
    const cachedData = cache.get(cacheKey);
    if (cachedData) return JSON.parse(cachedData);
    const ss = SpreadsheetApp.openById(SPREADSHEET_DASH);
    let nps = "-", venda = "-", bh = "-", quadro = "-", txdes = "-";
    const abasConsultadas = ['HISTORICO_VENDA_ANO', 'HISTORICO_NPS', 'HISTORICO_BH_ACUMULADO', 'HISTORICO_QUADRO', 'HISTORICO_TXDESL'];
    const blocosDeDados = {};

    abasConsultadas.forEach(nomeAba => {
      const sheet = ss.getSheetByName(nomeAba);
      blocosDeDados[nomeAba] = sheet ? sheet.getDataRange().getValues() : [];
    });

    const dVendas = blocosDeDados['HISTORICO_VENDA_ANO'];
    for(let i = 1; i < dVendas.length; i++) {
      let rowF = parseInt(dVendas[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let v = dVendas[i][7];
        if (typeof v === 'number') venda = (v * 100).toFixed(2);
        else { 
          let num = parseFloat(String(v).replace('%', '').replace(',', '.')); 
          if (!isNaN(num)) venda = (num < 10 ? num * 100 : num).toFixed(2); 
        }
        break;
      }
    }

    const dNPS = blocosDeDados['HISTORICO_NPS'];
    let maxDateNPS = 0; 
    let lastNPS = "-";
    for(let i = 1; i < dNPS.length; i++) {
      let rowF = parseInt(dNPS[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dNPS[i][0]);
        if (currentDt >= maxDateNPS) {
          maxDateNPS = currentDt;
          let v = parseFloat(String(dNPS[i][6]).replace(',', '.'));
          if(!isNaN(v)) lastNPS = v.toFixed(1);
        }
      }
    }
    nps = lastNPS;

    const dBH = blocosDeDados['HISTORICO_BH_ACUMULADO'];
    let maxDateBH = 0; 
    let sumBH = 0; 
    let matchFoundBH = false;
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(dBH[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt > maxDateBH) maxDateBH = currentDt;
      }
    }
    for(let i = 1; i < dBH.length; i++) {
      let rowF = parseInt(dBH[i][1], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let currentDt = obterDataRawSegura(dBH[i][0]);
        if (currentDt === maxDateBH) {
          let v = parseFloat(String(dBH[i][9]).replace(',', '.'));
          if(!isNaN(v)) { sumBH += v; matchFoundBH = true; }
        }
      }
    }
    if(matchFoundBH) bh = sumBH.toFixed(2);

    const dQuadro = blocosDeDados['HISTORICO_QUADRO'];
    let qContratar = 0; 
    let matchFoundQ = false;
    for(let i = 1; i < dQuadro.length; i++) {
      let rowF = parseInt(dQuadro[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let cargo = String(dQuadro[i][7]).trim();
        if (cargo !== "Intermitente" && cargo !== "Outros - Montagem") {
          let contratarVal = parseFloat(String(dQuadro[i][12]).replace(',', '.'));
          if(!isNaN(contratarVal)) { qContratar += contratarVal; matchFoundQ = true; }
        }
      }
    }
    if(matchFoundQ) quadro = Math.round(qContratar);

    const dTx = blocosDeDados['HISTORICO_TXDESL'];
    for(let i = 1; i < dTx.length; i++) {
      let rowF = parseInt(dTx[i][0], 10);
      if (rowF > 3000) rowF -= 3000;
      if (rowF === targetFilial) {
        let v = parseFloat(String(dTx[i][9]).replace('%', '').replace(',', '.'));
        if(!isNaN(v)) {
          let perc = (v * 100).toFixed(2);
          if (txdes === "-" || parseFloat(perc) > parseFloat(txdes)) txdes = perc;
        }
      }
    }
    let payload = {
      sucesso: true,
      venda: venda !== "-" ? String(venda).replace('.', ',') + "%" : "-",
      nps: nps !== "-" ? String(nps).replace('.', ',') : "-",
      bancoHoras: bh !== "-" ? String(bh).replace('.', ',') : "-",
      quadro: quadro !== "-" ? String(quadro) : "-",
      txdes: txdes !== "-" ? String(txdes).replace('.', ',') + "%" : "-"
    };
    cache.put(cacheKey, JSON.stringify(payload), 300);
    return payload;
  } catch(e) { return { erro: e.message }; }
}

// =============================================================================
// GRAVAÇÕES E ESCRITAS CORPORATIVAS (LOCKSERVICE BLINDADO)
// =============================================================================

function registrarAtividade(dados, fileData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      throw new Error("Usuário não autorizado a realizar lançamentos.");
    }
    registrarAuditoria("Nova Atividade", dados.motivo + " | " + dados.destino);
    try {
      const cache = CacheService.getScriptCache();
      if (dados.destino && !String(dados.destino).startsWith("REGIONAL") && !String(dados.destino).startsWith("DIRETORIA")) {
          let fId = parseInt(String(dados.destino).split('-')[0].trim(), 10);
          if (!isNaN(fId)) cache.remove('IND_LOJA_' + (fId > 3000 ? fId - 3000 : fId));
      }
    } catch(err) {}

    let linkEvidencia = "";
    if (fileData && fileData.base64) {
      let safeName = fileData.fileName || "evidencia_" + new Date().getTime() + ".png";
      let safeMime = fileData.mimeType || "image/png";
      linkEvidencia = DriveApp.getFolderById(EVIDENCIAS_FOLDER_ID).createFile(
        Utilities.newBlob(Utilities.base64Decode(fileData.base64), safeMime, safeName)
      ).getUrl();
    } else if (dados.linkEvidenciaExterna) {
      linkEvidencia = dados.linkEvidenciaExterna;
    }

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DADOS_LANCAMENTOS');
    const arrayInsert = new Array(27).fill("");
    const dHora = new Date();

    arrayInsert[0] = "ACT-" + dHora.getTime();
    arrayInsert[1] = dHora;
    arrayInsert[2] = emailLogado;
    arrayInsert[3] = dados.destino;
    arrayInsert[4] = dados.motivo;
    arrayInsert[5] = "";
    arrayInsert[6] = "";
    arrayInsert[7] = ""; // Removido score checklist
    arrayInsert[8] = "";
    arrayInsert[9] = (dados.moedas !== undefined) ? dados.moedas : 1;
    arrayInsert[10] = dados.observacoes || "";
    arrayInsert[11] = linkEvidencia;
    arrayInsert[12] = dados.dataViagem || dados.dataAtividade || "";
    arrayInsert[13] = "";
    arrayInsert[14] = dados.kmValor || "";
    arrayInsert[15] = dados.kmQtd || "";
    arrayInsert[16] = dados.kmCusto || "";
    arrayInsert[17] = dados.roteiro || "";
    arrayInsert[18] = dados.subTema || "";
    arrayInsert[19] = dados.pessoasImpactadas || 0;
    arrayInsert[20] = dados.tempoGasto || 0;
    arrayInsert[21] = dados.gastoTotal || 0;
    arrayInsert[22] = dados.alimentacao || 0;
    arrayInsert[23] = dados.hospedagem || 0;
    arrayInsert[24] = dados.aereo || 0;
    arrayInsert[25] = dados.pedagio || 0;
    arrayInsert[26] = dados.estacionamento || 0;
    sheet.appendRow(arrayInsert);
    return { sucesso: true };
  } catch (e) { 
    return { erro: e.message }; 
  } finally {
    lock.releaseLock();
  }
}

function excluirLancamento(idLancamento) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
    const controle = obterControleAcesso(emailLogado);
    if (!controle.autorizado) {
      throw new Error("Operação não autorizada.");
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('DADOS_LANCAMENTOS');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idLancamento) {
        if (String(data[i][2]).toLowerCase().trim() === emailLogado || controle.isSuperAdmin) {
            try {
                const destinoStr = String(data[i][3]).trim();
                const cache = CacheService.getScriptCache();
                if (!destinoStr.startsWith("REGIONAL") && !destinoStr.startsWith("DIRETORIA")) {
                    let fId = parseInt(destinoStr.split('-')[0].trim(), 10);
                    if (!isNaN(fId)) cache.remove('IND_LOJA_' + (fId > 3000 ? fId - 3000 : fId));
                }
            } catch(e) {}
            sheet.deleteRow(i + 1); 
            return { sucesso: true };
        }
      }
    }
    return { erro: "Documento não localizado ou permissões insuficientes." };
  } catch (e) { 
    return { erro: e.message }; 
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// GESTÃO DE NATUREZAS E AVISOS (RESTRIÇÃO ADMINISTRATIVA SERVER-SIDE)
// =============================================================================

function adicionarNatureza(novaNatureza) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Apenas administradores do sistema podem registrar naturezas." };
  }
  try {
    if(!novaNatureza) return { erro: "Nome inválido!" };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
    const valores = sheetConfig.getDataRange().getValues().map(r => String(r[0]).toLowerCase().trim());
    if (valores.includes(novaNatureza.toLowerCase().trim())) return { erro: "Esta natureza já está cadastrada!" };

    sheetConfig.appendRow([novaNatureza.trim()]);
    return { sucesso: true, naturezas: carregarNaturezasSeguras(ss) };
  } catch (e) { return { erro: e.message }; }
}

function excluirNatureza(naturezaTexto) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Operação restrita a administradores." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetConfig = ss.getSheetByName('CONFIGURAÇÕES');
    const valores = sheetConfig.getDataRange().getValues();
    for (let i = 1; i < valores.length; i++) {
      if (String(valores[i][0]).trim() === naturezaTexto.trim()) {
        sheetConfig.deleteRow(i + 1);
        return { sucesso: true, naturezas: carregarNaturezasSeguras(ss) };
      }
    }
    return { erro: "Natureza não localizada." };
  } catch (e) { return { erro: e.message }; }
}

function publicarAviso(mensagem, diretoriaAlvo) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Publicação restrita a diretores ou administradores." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    let nomeAutor = controle.nome;
    const usu = ss.getSheetByName('DADOS_USUARIOS').getDataRange().getValues();
    let emails = [];

    let dirAlvoNorm = String(diretoriaAlvo).toLowerCase().trim();
    for (let i = 1; i < usu.length; i++) {
       let em = String(usu[i][0]).toLowerCase().trim();
       let dirColNormArray = String(usu[i][3]).toLowerCase().split(',').map(s => s.trim());

       if(em && (dirAlvoNorm === "todos" || dirColNormArray.includes(dirAlvoNorm))) {
           emails.push(em);
       }
    }
    let prefixo = diretoriaAlvo === "Todos" ? "" : `[${diretoriaAlvo}] `;
    sheetAvisos.appendRow([new Date(), nomeAutor, prefixo + mensagem]);
    if (emails.length > 0) {
      MailApp.sendEmail({
        bcc: emails.join(','), subject: "📢 Novo Recado no Portal GP 360", 
        htmlBody: `<div style="font-family:Arial;border:1px solid #ddd;padding:20px;border-radius:8px;"><h2 style="color:#0086ff;">Novo Recado!</h2><p><strong>De:</strong> ${nomeAutor}</p><p><strong>Para:</strong> ${diretoriaAlvo}</p><div style="background:#fff3cd;padding:15px;border-left:5px solid #ffcc00;">"${mensagem}"</div><br><a href="${ScriptApp.getService().getUrl()}" style="background:#0086ff;color:white;padding:10px 20px;text-decoration:none;border-radius:5px;">Acessar Portal</a></div>`
      });
    }
    return { sucesso: true };
  } catch (e) { return { erro: e.message }; }
}

function excluirAvisoPlanilha(dataAviso, mensagemAviso) {
  const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
  const controle = obterControleAcesso(emailLogado);
  if (!controle.isSuperAdmin && !controle.isConfigAdmin) {
    return { erro: "ACESSO NEGADO: Apenas administradores podem excluir recados." };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetAvisos = ss.getSheetByName('DADOS_AVISOS');
    const dados = sheetAvisos.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
       if (formatarDataSegura(dados[i][0]) === dataAviso && String(dados[i][2]) === mensagemAviso) {
           sheetAvisos.deleteRow(i + 1); return { sucesso: true };
       }
    }
    return { erro: "Recado não localizado." };
  } catch (e) { return { erro: e.message }; }
}
