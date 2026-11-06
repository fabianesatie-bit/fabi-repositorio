//=============================================================================
// BACKEND OTIMIZADO: HUB DE GOVERNANÇA GP 360 (Visão Geral GP Sidebar)
//=============================================================================

const DB_MASTER = '1Nk0F5_tzevdbfmOTnpmhePdum6N22Ctf7g1N_ojuSjA'; // GP360
const DB_DASH   = '1FKcQtoGI5Hz8vYefD450EcnO8rW36sTSPIsAQkEIVlc'; // Indicadores Novos
const DB_APUR   = '1tn2FiNVWVMFM-3DC14_L0LaHGmd9O-teU2cvsIoajkk'; // Apurações
const DB_SOC    = '1InLKT3qmWxAv7N-U1tyoSW0tNI-Qc4LTW0vyza1oSg0'; // Social
const DB_AGENTE = '1Ldv8tzM0LucrhpntUwbB-SuIYuWlObjTShU6mh88j58'; // Agente

function doGet() {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle('GP 360 - Hub de Governança')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * RASTREADOR: Registra cliques do usuário na planilha de auditoria corporativa.
 * Protegido com LockService contra concorrência e execuções manuais fantasmas.
 */
function registrarLogSidebar(acao, detalhe) {
    if (!acao || !detalhe) {
        console.warn("Execução manual ignorada: Parâmetros 'acao' e 'detalhe' ausentes.");
        return; 
    }

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); 
        const ssId = '1phPQnIBiyVC1OqxooDQhyrR3_aR84jtqYnPJyOij0lY';
        const ss = SpreadsheetApp.openById(ssId);
        
        let sheet = ss.getSheetByName('AUDITORIA_SIDEBAR');
        if (!sheet) {
            sheet = ss.insertSheet('AUDITORIA_SIDEBAR');
            sheet.appendRow(["DATA_HORA", "USUARIO", "ACAO", "DETALHE"]);
            sheet.getRange("A1:D1").setFontWeight("bold");
        }
        
        const email = Session.getActiveUser().getEmail().toLowerCase().trim() || 'usuario_desconhecido';
        sheet.appendRow([new Date(), email, String(acao), String(detalhe)]);
        
    } catch(e) {
        console.error("Falha ao registrar log de telemetria: " + e.message);
    } finally {
        lock.releaseLock();
    }
}

// ============================================================================
// HELPERS DE PARSE E FORMATAÇÃO DE DADOS
// ============================================================================

function normalizarFilialId(id) {
    if (!id) return null;
    let num = parseInt(String(id).replace(/\D/g, ''), 10);
    if (isNaN(num)) return null;
    if (num > 3000) num -= 3000; 
    return num;
}

function extrairIdDrive(linkOuId) {
    if (!linkOuId) return "";
    const match = String(linkOuId).match(/[-\w]{25,}(?!.*[-\w]{25,})/);
    return match ? "https://lh3.googleusercontent.com/d/" + match[0] : (String(linkOuId).startsWith("http") ? linkOuId : "");
}

function normalizeName(str) {
    return String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

/**
 * NORMALIZADOR DE HORAS (Anti-Timestamp de Objeto Date do Google Sheets)
 * Converte Date, strings "HH:MM", decimais e inteiros em valor arredondado de no máximo 2 dígitos.
 */
function normalizarTempoHoras(val) {
    if (val === null || val === undefined || val === '') return 0;
    let horas = 0;

    // Se o Google Sheets entregou um objeto Date (horário como 04:00:00 ou fórmula)
    if (val instanceof Date) {
        horas = val.getHours() + (val.getMinutes() / 60);
    } else if (typeof val === 'string') {
        let str = val.trim().replace(',', '.');
        let matchHora = str.match(/^(\d{1,2}):(\d{2})/);
        if (matchHora) {
            horas = parseInt(matchHora[1], 10) + (parseInt(matchHora[2], 10) / 60);
        } else {
            let n = parseFloat(str.replace(/[^0-9.-]/g, ''));
            horas = isNaN(n) ? 0 : n;
        }
    } else if (typeof val === 'number') {
        horas = val;
    }

    // Se for valor negativo ou timestamp espúrio em milissegundos (> 1000 horas em uma única ação)
    if (horas < 0 || horas > 1000) return 0;

    // Arredonda para número inteiro (máximo 2 dígitos por ação individual, ex: até 99h)
    let arredondado = Math.round(horas);
    return arredondado > 99 ? 99 : arredondado;
}

/**
 * BLINDAGEM ANTI-FUSO HORÁRIO (Timezone Shift) E TIMESTAMPS
 */
function dataParaStringISO(dataVal, fallbackIsoData = null) {
    if (!dataVal || String(dataVal).trim() === "") return fallbackIsoData;
    try {
        if (dataVal instanceof Date) {
            return Utilities.formatDate(dataVal, "America/Sao_Paulo", "yyyy-MM-dd");
        }

        let str = String(dataVal).trim();
        
        // Remove hora se houver (espaço ou T)
        if (str.includes("T")) str = str.split("T")[0].trim();
        if (str.includes(" ")) str = str.split(" ")[0].trim();

        if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
            return str.substring(0, 10);
        }

        let cleanStr = str.replace(/Horário Padrão de Brasília/gi, '').replace(/GMT.*/, '').trim();
        
        let matchBR = cleanStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (matchBR) {
            let dia = matchBR[1].padStart(2, '0');
            let mes = matchBR[2].padStart(2, '0');
            let ano = matchBR[3];
            return `${ano}-${mes}-${dia}`;
        }

        let dt = new Date(cleanStr);
        if (!isNaN(dt.getTime())) {
            let ano = dt.getFullYear();
            let m = (dt.getMonth() + 1).toString().padStart(2, '0');
            let d = dt.getDate().toString().padStart(2, '0');
            return `${ano}-${m}-${d}`;
        }
        return fallbackIsoData;
    } catch(e) { return fallbackIsoData; }
}

function extrairDataBR(dataValor) {
    if (!dataValor || String(dataValor).trim() === "") return "-";
    try {
        let str = String(dataValor).trim();
        if (str.match(/ok/i) || str.match(/pendente/i) || str.match(/respondido/i) || str.length < 6) return str; 
        let iso = dataParaStringISO(dataValor);
        if(iso) { let p = iso.split('-'); return `${p[2]}/${p[1]}/${p[0]}`; }
        return str; 
    } catch(e) { return "-"; }
}

function extrairMesAno(dataValor) {
    let iso = dataParaStringISO(dataValor);
    if(iso) return iso.substring(0, 7);
    return "0000-00";
}

function obterMapaAbas(ss) {
    let mapa = {};
    if (!ss) return mapa;
    let sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
        mapa[normalizeName(sheets[i].getName())] = sheets[i];
    }
    return mapa;
}

function buscarAbaNoMapa(mapaAbas, names) {
    if (!mapaAbas) return null;
    let targetNames = names.map(normalizeName);
    for (let i = 0; i < targetNames.length; i++) {
        if (mapaAbas[targetNames[i]]) return mapaAbas[targetNames[i]];
    }
    // Fallback flexível ignorando pontuações e espaços
    let keys = Object.keys(mapaAbas);
    let targetCleans = names.map(n => normalizeName(n).replace(/[^a-z0-9]/g, ''));
    for (let i = 0; i < targetCleans.length; i++) {
        let tc = targetCleans[i];
        for (let j = 0; j < keys.length; j++) {
            let kClean = keys[j].replace(/[^a-z0-9]/g, '');
            if (kClean === tc || kClean.includes(tc) || tc.includes(kClean)) {
                return mapaAbas[keys[j]];
            }
        }
    }
    return null;
}

/**
 * Localização de aba de alta performance: tenta getSheetByName primeiro (10ms)
 * antes de listar todas as abas com getSheets() (3000ms).
 */
function getSheetFast(ss, names) {
    if (!ss) return null;
    for (let i = 0; i < names.length; i++) {
        let sh = ss.getSheetByName(names[i]);
        if (sh) return sh;
    }
    let sheets = ss.getSheets();
    let cleanTargets = names.map(n => normalizeName(n).replace(/[^a-z0-9]/g, ''));
    for (let i = 0; i < sheets.length; i++) {
        let sClean = normalizeName(sheets[i].getName()).replace(/[^a-z0-9]/g, '');
        for (let j = 0; j < cleanTargets.length; j++) {
            if (sClean === cleanTargets[j] || sClean.includes(cleanTargets[j]) || cleanTargets[j].includes(sClean)) {
                return sheets[i];
            }
        }
    }
    return null;
}

function obterValoresSeguros(sh) {
    if (!sh) return [];
    let lr = sh.getLastRow();
    let lc = sh.getLastColumn();
    if (lr < 1 || lc < 1) return [];
    return sh.getRange(1, 1, lr, lc).getValues();
}

function getSheetSafe(ss, names) {
    return getSheetFast(ss, names);
}

// ============================================================================
// DIAGNÓSTICO RÁPIDO DE CONECTIVIDADE (Execute direto no editor se houver lentidão)
// ============================================================================

function testarConexoesBases() {
    const bases = [
        { nome: 'DB_MASTER (GP360)', id: DB_MASTER },
        { nome: 'DB_DASH (Indicadores Novos)', id: DB_DASH },
        { nome: 'DB_APUR (Apurações)', id: DB_APUR },
        { nome: 'DB_SOC (Social)', id: DB_SOC },
        { nome: 'DB_AGENTE (Agente)', id: DB_AGENTE }
    ];
    
    console.log("=== INICIANDO DIAGNÓSTICO DE CONEXÃO DAS BASES ===");
    bases.forEach(b => {
        let t0 = new Date().getTime();
        try {
            let ss = SpreadsheetApp.openById(b.id);
            let tOpen = new Date().getTime() - t0;
            let sheets = ss.getSheets().map(s => `${s.getName()} (${s.getLastRow()} lin x ${s.getLastColumn()} col)`);
            console.log(`✅ [${b.nome}] Acessada em ${tOpen}ms | Abas: ${sheets.join(' | ')}`);
        } catch(e) {
            let tFail = new Date().getTime() - t0;
            console.error(`❌ [${b.nome}] FALHA após ${tFail}ms: ${e.message}`);
        }
    });
    console.log("=== DIAGNÓSTICO FINALIZADO ===");
}

function testarAbaDesligamentos() {
    console.log("=== TESTE ESPECÍFICO: ABA DESLIGAMENTOS ===");
    try {
        let ssApur = SpreadsheetApp.openById(DB_APUR);
        let sh = getSheetFast(ssApur, [
            'HISTORICO_DESLIGAMENTO_F', 'Historico_Desligamento_F', 'historico_desligamento_f',
            'Historico_Desligamentos', 'HISTORICO_DESLIGAMENTOS', 
            'Historico Desligamentos', 'Histórico de Desligamentos', 
            'Desligamentos', 'DESLIGAMENTOS'
        ]);
        if (!sh) {
            console.error("❌ Nenhuma aba de Desligamentos encontrada em DB_APUR!");
            console.log("Abas existentes: " + ssApur.getSheets().map(s => s.getName()).join(', '));
            return;
        }
        console.log(`✅ Aba de Desligamentos encontrada: "${sh.getName()}" (${sh.getLastRow()} linhas x ${sh.getLastColumn()} colunas)`);
        let d = obterValoresSeguros(sh);
        if (d.length > 0) {
            console.log("Cabeçalho (colunas):", d[0].map((c, i) => `[Col ${i}] ${c}`).join(' | '));
            if (d.length > 1) {
                console.log("Exemplo Linha 1:", d[1].map((c, i) => `[Col ${i}] ${c}`).join(' | '));
            }
        }
    } catch(e) {
        console.error("❌ Erro ao inspecionar aba desligamentos:", e.message);
    }
}

// ============================================================================
// SERVIÇOS DE EXPORTAÇÃO
// ============================================================================

function gerarPlanilhaStatus(dadosJson) {
    try {
        let dados = JSON.parse(dadosJson);
        let timestamp = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd-MM-yyyy_HHmm");
        let ss = SpreadsheetApp.create("Status_Regionais_GP360_" + timestamp);
        let sh = ss.getActiveSheet();
        
        sh.appendRow(["Filial", "Regional", "Vendas Acum. (Ano)", "NPS (Loja)", "Banco Hrs", "Tx Desl", "Vagas Totais", "Visitas In Loco"]);
        
        let matrix = dados.map(d => [d.filial, d.regional, d.vendas, d.nps, d.banco, d.tx, d.vagas, d.visitas]);
        if(matrix.length > 0) {
            sh.getRange(2, 1, matrix.length, 8).setValues(matrix);
        }
        
        let header = sh.getRange("A1:H1");
        header.setFontWeight("bold").setBackground("#0086FF").setFontColor("#FFFFFF");
        sh.setFrozenRows(1);
        sh.autoResizeColumns(1, 8);
        
        return JSON.stringify({sucesso: true, url: ss.getUrl()});
    } catch(e) {
        return JSON.stringify({sucesso: false, erro: e.message});
    }
}

// ============================================================================
// MOTOR PRINCIPAL DE DADOS (Consolidação de Lojas, Ações e KPIs)
// ============================================================================

function obterCargaMestreDashboard() {
    let avisos = [];
    console.time("CargaMestre_Total");
    try {
        let emailLogado = "";
        try {
            emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
        } catch(eMailErr) {
            console.warn("Não foi possível capturar o e-mail ativo:", eMailErr.message);
        }
        
        console.log("Iniciando obterCargaMestreDashboard para o usuário:", emailLogado || "(E-mail não capturado)");

        // LISTA DE ADMINISTRADORES AUTORIZADOS
        const administradores = [
            'fabiane.satie@magazineluiza.com.br',
            'lorena.castro@magazineluiza.com.br'
        ];
        const isAdmin = administradores.includes(emailLogado);
        
        // PERMISSÃO DE ROI
        const canViewRoi = isAdmin || (emailLogado === 'tarcisio.maniglia@magazineluiza.com.br');
        
        let regionaisPermitidas = new Set();
        let role = isAdmin ? 'ADMIN' : 'COORDENADOR';
        let acesso = { email: emailLogado, isAdmin: isAdmin, canViewRoi: canViewRoi, role: role };
        
        let masterData = {}; 
        let acoesMuralGlobais = [];
        let visitTracker = new Set(); // Engine Universal de Deduplicação
        
        let filtrosGlobais = { meses: new Set(), diretorias: new Set(), regionais: new Set(), filiais: new Set(), coordenadores: new Set() };
        let coordFotos = {};
        let emailToName = {};

        let hojeIsoObj = new Date();
        let hojeIsoDia = Utilities.formatDate(hojeIsoObj, "America/Sao_Paulo", "yyyy-MM-dd");
        let hojeIsoMes = Utilities.formatDate(hojeIsoObj, "America/Sao_Paulo", "yyyy-MM");

        let ssMaster;
        try {
            console.time("Leitura_DB_MASTER");
            ssMaster = SpreadsheetApp.openById(DB_MASTER);
        } catch(e) {
            return JSON.stringify({ sucesso: false, erro: `Falha ao acessar GP360. Permissão negada no Drive para o e-mail: ${emailLogado || 'desconhecido'}` });
        }
        
        let injetarAcaoVirtual = (idFilial, dataRaw, motivo, obs, peso, origem = 'VIRTUAL', autorEspecifico = null) => {
            if(!idFilial || !masterData[idFilial]) return;
            let m = extrairMesAno(dataRaw);
            if (m < '2023-01' && m !== "0000-00") return; // Ignora registros antigos no mural
            let autor = autorEspecifico || masterData[idFilial].coordenadores[0] || 'Sistema';
            acoesMuralGlobais.push({
                idFilial: idFilial, destinoBruto: idFilial, mes: m, dataStr: extrairDataBR(dataRaw), dataIso: dataParaStringISO(dataRaw),
                autor: autor, motivo: motivo, obs: obs,
                gastos: 0, km: 0, moedas: peso, pessoas: 1, tempo: 1,
                isVisita: false, isVisitaUnica: false, isSocial: motivo.toUpperCase().includes('SOCIAL'),
                origemVirtual: origem, isAvulsoFinanceiro: false
            });
        };

        const shUsu = ssMaster.getSheetByName('DADOS_USUARIOS');
        let coordMap = {}; 

        if (shUsu) {
            const dadosUsu = obterValoresSeguros(shUsu);
            for (let i = 1; i < dadosUsu.length; i++) {
                let em = String(dadosUsu[i][0] || '').toLowerCase().trim();
                let nome = String(dadosUsu[i][1] || '').trim();
                let cargo = String(dadosUsu[i][2] || '').toUpperCase();
                let regStr = String(dadosUsu[i][4] || '').toUpperCase();
                let nivelAcesso = String(dadosUsu[i][5] || '').toUpperCase();
                let fotoUrl = extrairIdDrive(dadosUsu[i][6]);
                
                if (nivelAcesso.includes('SEM ACESSO') || (cargo.includes('REGIONAL OP') && !cargo.includes('GP')) || cargo === 'GERENTE DE LOJA') continue;

                if (em) emailToName[em] = nome;
                if (nome && fotoUrl) coordFotos[nome] = fotoUrl;

                if (em === emailLogado) {
                    if (cargo.includes('GERENTEGP') || cargo.includes('GERENTE GP')) acesso.role = 'GERENTEGP';
                    else if (cargo.includes('DIRETOR') || cargo.includes('DIRETORIA')) acesso.role = 'DIRETORRH';
                    else if (isAdmin || nivelAcesso.includes('MASTER') || nivelAcesso.includes('ADMIN')) acesso.role = 'ADMIN';
                    else acesso.role = 'COORDENADOR';
                    
                    if (regStr) regStr.split(',').forEach(r => regionaisPermitidas.add(r.trim().toUpperCase()));
                }

                if (nome && (cargo.includes('COORDENAD') || nivelAcesso.includes('COORDENAD'))) {
                    filtrosGlobais.coordenadores.add(nome);
                    if (regStr) {
                        regStr.split(',').forEach(r => {
                            let rLimpa = r.trim().toUpperCase();
                            if(!coordMap[rLimpa]) coordMap[rLimpa] = [];
                            if(!coordMap[rLimpa].includes(nome)) coordMap[rLimpa].push(nome);
                        });
                    } else {
                        if(!coordMap['SEM_REGIONAL']) coordMap['SEM_REGIONAL'] = [];
                        coordMap['SEM_REGIONAL'].push(nome);
                    }
                }
            }
        }
        
        // Se for admin, garante acesso irrestrito
        if (isAdmin) {
            regionaisPermitidas.add('TODAS');
        } else if (regionaisPermitidas.size === 0 && acesso.role !== 'DIRETORRH' && acesso.role !== 'GERENTEGP') {
            return JSON.stringify({ sucesso: false, erro: `Conta (${emailLogado || 'desconhecida'}) sem regionais autorizadas. Solicite liberação na aba DADOS_USUARIOS.` });
        }

        const shLojas = ssMaster.getSheetByName('DADOS_LOJAS');
        if (!shLojas) return JSON.stringify({ sucesso: false, erro: "Aba DADOS_LOJAS não encontrada." });
        
        const dadosLojas = obterValoresSeguros(shLojas);
        for (let i = 1; i < dadosLojas.length; i++) {
            let id = normalizarFilialId(dadosLojas[i][0]);
            let nome = String(dadosLojas[i][1] || '').trim();
            let regional = String(dadosLojas[i][2] || '').toUpperCase().trim();
            let diretoria = String(dadosLojas[i][3] || '').toUpperCase().trim();

            if (id && (isAdmin || regionaisPermitidas.has(regional) || regionaisPermitidas.has('TODAS') || acesso.role === 'DIRETORRH' || acesso.role === 'GERENTEGP')) {
                let coordsAtendendo = coordMap[regional] ? [...coordMap[regional]] : [];
                if (coordMap['TODAS']) coordMap['TODAS'].forEach(c => { if (!coordsAtendendo.includes(c)) coordsAtendendo.push(c); });
                if (coordMap['SEM_REGIONAL']) coordMap['SEM_REGIONAL'].forEach(c => { if (!coordsAtendendo.includes(c)) coordsAtendendo.push(c); });

                masterData[id] = {
                    id: id, nome: nome, regional: regional, diretoria: diretoria, coordenadores: coordsAtendendo,
                    historicoGP: [], social: [],
                    agente: { elegiveis: 0, concluidos: 0 },
                    indicadores: { venda: '-', nps: '-', bancoHoras: 0, txSaida: '-' },
                    pendencias: {
                        apuracoes: [], rascunhos: [], comite: [], feedbacksSLA: [], desligamentos: [], intervencoes: [],
                        agenteFeedbacks: [], bancoHoras: [], socialTodos: [], vagasInfo: [], npsVendedores: []
                    }
                };
                filtrosGlobais.diretorias.add(diretoria);
                filtrosGlobais.regionais.add(regional);
                filtrosGlobais.filiais.add(id.toString());
            }
        }

        const shLanc = ssMaster.getSheetByName('DADOS_LANCAMENTOS');
        if (shLanc) {
            const dadosLanc = obterValoresSeguros(shLanc);
            console.log("DADOS_LANCAMENTOS carregados:", dadosLanc.length, "linhas.");

            for(let i = 1; i < dadosLanc.length; i++) {
                let dataAcao = dadosLanc[i][12]; 
                if (!dataAcao) continue;

                let destinoFilialBruto = String(dadosLanc[i][3] || '').trim(); 
                let idL = normalizarFilialId(destinoFilialBruto); 

                let motivoVal = String(dadosLanc[i][4] || '').trim(); 
                let tipoRoteiro = String(dadosLanc[i][17] || ''); 
                
                let rotUpper = String(tipoRoteiro || '').toUpperCase();
                let motivoUpper = motivoVal.toUpperCase();
                let isVisita = rotUpper.includes("PRESENCIAL") || rotUpper.includes("IN LOCO") || rotUpper.includes("INLOCO") || rotUpper.includes("VISITA");

                // BARREIRA FINANCEIRA AVULSO
                let isAvulso = destinoFilialBruto.toUpperCase().startsWith('AVULSO') || motivoUpper.includes('AVULSO') || motivoUpper.includes('SEDE/EVENTO');
                let isSocial = motivoUpper.includes('SOCIAL') || motivoUpper.includes('ASSIST') || rotUpper.includes('SOCIAL') || rotUpper.includes('ASSIST');

                let mesStr = extrairMesAno(dataAcao);
                if (mesStr < '2023-01' && mesStr !== "0000-00") continue; // Pula lançamentos muito antigos para otimizar velocidade
                let isoAcao = dataParaStringISO(dataAcao);
                
                let emA = String(dadosLanc[i][2] || '').toLowerCase().trim();
                let nomeAutorCorreto = emailToName[emA] || emA.split('@')[0];

                // ENGINE DE DEDUPLICAÇÃO DE VISITAS
                let isVisitaUnica = false;
                if (isVisita && isoAcao && !isAvulso) {
                    let destKey = idL ? idL : destinoFilialBruto.toUpperCase().replace(/\s/g, '');
                    let chvVisita = `${emA}_${destKey}_${isoAcao}`;
                    if (!visitTracker.has(chvVisita)) {
                        visitTracker.add(chvVisita);
                        isVisitaUnica = true;
                    }
                }

                let ehRegionalAtendida = false;
                if (idL && masterData[idL]) ehRegionalAtendida = true;
                if (!idL && (isAdmin || regionaisPermitidas.has('TODAS'))) ehRegionalAtendida = true;
                if (!idL && !ehRegionalAtendida) {
                    for (let regPermitida of regionaisPermitidas) {
                        if (destinoFilialBruto.toUpperCase().includes(regPermitida)) { ehRegionalAtendida = true; break; }
                    }
                }

                let passaFiltroAcesso = false;
                if (acesso.role === 'ADMIN' || acesso.role === 'GERENTEGP' || acesso.role === 'DIRETORRH') {
                    if (ehRegionalAtendida) passaFiltroAcesso = true;
                } else {
                    if (emA === emailLogado) passaFiltroAcesso = true;
                }

                if (passaFiltroAcesso) {
                    if(mesStr !== "0000-00") filtrosGlobais.meses.add(mesStr);
                    
                    // NORMALIZAÇÃO DE TEMPO (HORAS) COM PROTEÇÃO ANTI-TIMESTAMP
                    let tempoHoras = isAvulso ? 0 : normalizarTempoHoras(dadosLanc[i][20]);

                    let acaoObj = {
                        idFilial: idL, destinoBruto: destinoFilialBruto, mes: mesStr, dataStr: extrairDataBR(dataAcao), dataIso: isoAcao, autor: nomeAutorCorreto,
                        motivo: motivoVal, obs: String(dadosLanc[i][10] || '').trim(),
                        gastos: parseFloat(String(dadosLanc[i][21] || '').replace(',','.')) || 0, 
                        km: parseFloat(String(dadosLanc[i][15] || '').replace(',','.')) || 0, 
                        
                        // Zera o peso se for apenas deslocamento avulso
                        moedas: isAvulso ? 0 : (parseFloat(String(dadosLanc[i][9] || '').replace(/\./g, '').replace(',','.')) || 0), 
                        pessoas: isAvulso ? 0 : (Number(dadosLanc[i][19]) || 0), 
                        tempo: tempoHoras,
                        
                        isVisita: isVisita && !isAvulso, isVisitaUnica: isVisitaUnica, isSocial: isSocial && !isAvulso, roteiro: tipoRoteiro, origemVirtual: 'FISICA',
                        isAvulsoFinanceiro: isAvulso
                    };

                    acoesMuralGlobais.push(acaoObj);
                }
            }
        }
        console.timeEnd("Leitura_DB_MASTER");

        // CONFIGURAÇÃO DB_DASH: Desativada temporariamente (dados não alimentados / Visão Geral removida)
        const ATIVAR_DB_DASH = false;

        if (ATIVAR_DB_DASH) {
            try {
                console.time("Leitura_DB_DASH");
                const ssDash = SpreadsheetApp.openById(DB_DASH);
                const mapaDash = obterMapaAbas(ssDash);
                
                let mapDash = (shNames, parser) => {
                    let sh = buscarAbaNoMapa(mapaDash, shNames);
                    if(sh) {
                        let d = obterValoresSeguros(sh);
                        for(let i=1; i<d.length; i++) parser(d[i]);
                    }
                };
                
                mapDash(['HISTORICO_QUADRO'], d => {
                    let idL = normalizarFilialId(d[0]);
                    let cargo = String(d[7] || '').trim();
                    let cargoUpper = cargo.toUpperCase();
                    if(idL && masterData[idL] && !cargoUpper.includes('INTERMITENTE') && !cargoUpper.includes('MONTAGEM') && !cargoUpper.includes('MONTADOR')) {
                        let v = parseFloat(d[12]);
                        if(!isNaN(v) && v > 0) { masterData[idL].pendencias.vagasInfo.push({ cargo: cargo, qtd: v }); }
                    }
                });
                
                let mapBHAux = {};
                mapDash(['HISTORICO_BH_ACUMULADO'], d => {
                    let idL = normalizarFilialId(d[1]);
                    let dtIso = dataParaStringISO(d[0], hojeIsoDia);
                    if(idL && masterData[idL] && dtIso <= hojeIsoDia) { 
                        let hrs = parseFloat(String(d[9]).replace(',','.'));
                        if(!isNaN(hrs)) {
                            if(!mapBHAux[idL]) mapBHAux[idL] = { maxDate: dtIso, total: 0, criticos: [] };
                            if(dtIso > mapBHAux[idL].maxDate) { mapBHAux[idL].maxDate = dtIso; mapBHAux[idL].total = 0; mapBHAux[idL].criticos = []; }
                            if(dtIso === mapBHAux[idL].maxDate) {
                                mapBHAux[idL].total += hrs;
                                if(hrs >= 5) mapBHAux[idL].criticos.push({ nome: String(d[6]).trim(), id: String(d[5]).trim(), horas: hrs });
                            }
                        }
                    }
                });
                Object.keys(mapBHAux).forEach(idL => {
                    masterData[idL].indicadores.bancoHoras = mapBHAux[idL].total;
                    masterData[idL].pendencias.bancoHoras = mapBHAux[idL].criticos;
                });

                let registrosNPSV = {}; 
                mapDash(['HISTORICO_NPS_VENDEDORES'], d => {
                    let nome = String(d[1]).trim();
                    let idL = normalizarFilialId(d[2]);
                    let dtIso = dataParaStringISO(d[0], hojeIsoDia);
                    let npsV = parseFloat(String(d[7]).replace(',','.'));
                    if(idL && nome && masterData[idL] && !isNaN(npsV) && dtIso <= hojeIsoDia) {
                        if(!registrosNPSV[idL]) registrosNPSV[idL] = { maxDate: dtIso, records: [] };
                        if(dtIso > registrosNPSV[idL].maxDate) { registrosNPSV[idL].maxDate = dtIso; registrosNPSV[idL].records = []; }
                        if(dtIso === registrosNPSV[idL].maxDate) { registrosNPSV[idL].records.push({ nome: nome, nota: npsV }); }
                    }
                });
                Object.keys(registrosNPSV).forEach(idL => {
                    registrosNPSV[idL].records.forEach(r => {
                        if(r.nota < 70) masterData[idL].pendencias.npsVendedores.push({ nome: r.nome, nota: r.nota });
                    });
                });

                console.timeEnd("Leitura_DB_DASH");

            } catch(e) {
                avisos.push("Base Dash inacessível: " + e.message);
                console.warn("Aviso DB_DASH:", e.message);
            }
        } else {
            console.log("DB_DASH ignorada (Visão Geral Desativada / Dados não alimentados)");
        }

        try {
            console.time("Leitura_DB_APUR");
            const ssApur = SpreadsheetApp.openById(DB_APUR);
            
            let mapApur = (shNames, parser) => {
                let sh = getSheetFast(ssApur, shNames);
                if(sh) {
                    let d = obterValoresSeguros(sh);
                    for(let i=1; i<d.length; i++) {
                        try { parser(d[i]); } catch(e){}
                    }
                }
            };

            mapApur(['HISTORICO_APURACAO', 'historico_apuracao'], r => {
                let idL = normalizarFilialId(r[3]); 
                let concl = String(r[7] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    let m = extrairMesAno(r[1]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                    let d1 = String(r[23]||''); let d2 = String(r[26]||''); let d3 = String(r[29]||'');
                    let alvoStr = [d1, d2, d3].filter(Boolean).join(', ');
                    masterData[idL].pendencias.apuracoes.push({ 
                        mes: m, data: extrairDataBR(r[1]), alvo: alvoStr, conclusao: String(r[7]||''), 
                        agendamento: extrairDataBR(r[20]==='sim'?r[21]:'N/A'), linkDoc: String(r[8]||'')
                    });
                    
                    let origemApur = String(r[2] || '').toUpperCase().includes('INTERNA') ? 'Apuração (Interna)' : 'Apuração (Canal)';
                    injetarAcaoVirtual(idL, r[1], origemApur, 'Processo de Apuração de Denúncia', 1, 'APURACAO');
                }
            });

            mapApur(['Rascunhos_Apuracoes', 'rascunhos_apuracoes'], r => {
                let idL = normalizarFilialId(r[3]); 
                let concl = String(r[7] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    let m = extrairMesAno(r[1]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                    let d1 = String(r[22]||''); let d2 = String(r[25]||''); let d3 = String(r[28]||'');
                    let alvoStr = [d1, d2, d3].filter(Boolean).join(', ');
                    let obj = { mes: m, data: extrairDataBR(r[1]), alvo: alvoStr, conclusao: String(r[7]||''), agendamento: '-', linkDoc: String(r[8]||'') };
                    if(concl.includes('COMIT')) { masterData[idL].pendencias.comite.push(obj); }
                    else { masterData[idL].pendencias.rascunhos.push(obj); }
                }
            });

            mapApur(['Intervencoes_Feedback', 'intervencoes_feedback'], r => {
                let idL = normalizarFilialId(r[1]); let concl = String(r[3] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    masterData[idL].pendencias.intervencoes.push({ data: extrairDataBR(r[0]), linkDoc: String(r[2]||''), status: String(r[3]||''), detalhe: String(r[4]||''), nHumor: String(r[5]||'-'), nData: extrairDataBR(r[6]) });
                    injetarAcaoVirtual(idL, r[0], 'Acompanhamento Feedbacks', 'Intervenção PAI', 1, 'INTERVENCAO');
                }
            });

            mapApur([
                'HISTORICO_DESLIGAMENTO_F', 'Historico_Desligamento_F', 'historico_desligamento_f',
                'HISTORICO_DESLIGAMENTOS', 'Historico_Desligamentos', 'Historico Desligamentos', 
                'Histórico de Desligamentos', 'Desligamentos', 'DESLIGAMENTOS'
            ], r => {
                let idL = normalizarFilialId(r[2]); // Coluna C: Filial
                if (!idL) return;

                let dataRaw = r[1]; // Coluna B: Data Registro
                let m = extrairMesAno(dataRaw);
                if (m !== "0000-00") filtrosGlobais.meses.add(m);

                let idDeslig = String(r[0] || '').trim(); // Coluna A: ID Desligamento (ex: DES-7F7892A0)
                let colabNome = String(r[3] || '').trim(); // Coluna D: Colaborador
                let matricula = String(r[4] || '').trim(); // Coluna E: ID / Matrícula
                let tempoEmp = String(r[5] || '').trim();  // Coluna F: Tempo Empresa
                let cargo = String(r[6] || '').trim();     // Coluna G: Cargo
                let resultados = String(r[7] || '').trim(); // Coluna H: Resultados
                let justificativa = String(r[8] || '').trim(); // Coluna I: Justificativa
                let evidenciasLink = String(r[9] || '').trim(); // Coluna J: Evidencias (Drive)
                let parecerCoord = String(r[10] || '').trim(); // Coluna K: Parecer Coordenador
                let stReg = String(r[12] || '').trim();   // Coluna M: Status Regional
                let parecerReg = String(r[13] || '').trim(); // Coluna N: Parecer Regional
                let stDir = String(r[15] || '').trim();   // Coluna P: Status Diretor Op
                let parecerDir = String(r[16] || '').trim(); // Coluna Q: Parecer Diretor Op
                let coordNome = String(r[18] || '').trim(); // Coluna S: Coordenador Nome
                let linkApuracao = String(r[20] || '').trim(); // Coluna U: Link Apuração (Google Docs)

                let parecerGeral = `${stReg} ${stDir} ${parecerReg} ${parecerDir}`.toUpperCase();
                if (parecerGeral.includes('CANCELADO') || parecerGeral.includes('LANÇADO INCORRETAMENTE') || parecerGeral.includes('LANCADO INCORRETAMENTE')) {
                    return;
                }

                // Identifica se o processo está em aberto (pendente de aprovação regional ou do diretor)
                let sRegU = stReg.toUpperCase();
                let sDirU = stDir.toUpperCase();
                let regConcluido = sRegU.includes('CONCLU') || sRegU.includes('APROV') || sRegU.includes('FINALIZ');
                let dirConcluido = sDirU.includes('CONCLU') || sDirU.includes('APROV') || sDirU.includes('FINALIZ');
                let isAberto = !(regConcluido && dirConcluido);

                // Dossiê principal: Coluna U (Link Apuração Docs), com fallback para Coluna J (Evidências Drive)
                let linkDoc = linkApuracao || evidenciasLink;

                // Normaliza link para garantir clique direto
                if (linkDoc && !linkDoc.startsWith('http')) {
                    if (linkDoc.includes('docs.google') || linkDoc.includes('drive.google')) {
                        linkDoc = 'https://' + linkDoc;
                    } else if (linkDoc.length >= 25 && !linkDoc.includes(' ')) {
                        linkDoc = 'https://docs.google.com/document/d/' + linkDoc + '/view';
                    }
                }

                if (!masterData[idL]) {
                    masterData[idL] = {
                        id: idL, nome: 'Filial ' + idL, regional: 'OUTROS', diretoria: 'OUTROS', coordenadores: [],
                        historicoGP: [], social: [], agente: { elegiveis: 0, concluidos: 0 },
                        indicadores: { venda: '-', nps: '-', bancoHoras: 0, txSaida: '-' },
                        pendencias: { apuracoes: [], rascunhos: [], comite: [], feedbacksSLA: [], desligamentos: [], intervencoes: [], agenteFeedbacks: [], bancoHoras: [], socialTodos: [], vagasInfo: [], npsVendedores: [] }
                    };
                    filtrosGlobais.filiais.add(idL.toString());
                }

                masterData[idL].pendencias.desligamentos.push({ 
                    idDeslig: idDeslig,
                    mes: m, 
                    data: extrairDataBR(dataRaw), 
                    colab: colabNome || 'Colaborador', 
                    matricula: matricula,
                    tempoEmp: tempoEmp,
                    cargo: cargo,
                    reg: stReg || 'Pendente Regional', 
                    dir: stDir || 'Pendente Diretor Op', 
                    coord: coordNome,
                    justificativa: justificativa,
                    dossieLink: linkDoc,
                    evidenciasLink: evidenciasLink,
                    isAberto: isAberto
                });
                
                injetarAcaoVirtual(idL, dataRaw, 'Relatório de Desligamento', 'Dossiê de Desligamento: ' + colabNome, 1, 'DESLIGAMENTO', coordNome);
            });

            mapApur(['Feedbacks_Gerentes', 'feedbacks_gerentes'], r => {
                let idL = normalizarFilialId(r[1]); let concl = String(r[5] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    let m = extrairMesAno(r[2]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                    masterData[idL].pendencias.feedbacksSLA.push({ mes: m, data: extrairDataBR(r[2]), gerente: String(r[9]||'').trim(), status: String(r[5]||''), dataRetorno: extrairDataBR(r[6]), evidenciaLink: String(r[8]||'') });
                }
            });
            console.timeEnd("Leitura_DB_APUR");
        } catch(e) {
            avisos.push("Base Apur inacessível: " + e.message);
            console.warn("Aviso DB_APUR:", e.message);
        }

        try {
            console.time("Leitura_DB_AGENTE");
            const ssAgente = SpreadsheetApp.openById(DB_AGENTE);
            
            let dSet = new Set();
            try {
                let shDesl = getSheetFast(ssAgente, ['atualizacao_desligamento', 'desligamentos', 'DESLIGAMENTOS']);
                if (shDesl) { 
                    let d = obterValoresSeguros(shDesl); 
                    for(let i=1; i<d.length; i++) if(d[i][0]) dSet.add(String(d[i][0]).trim()); 
                }
            } catch(e){}

            try {
                let shAgenteHis = getSheetFast(ssAgente, ['REGISTRO_FEEDBACK', 'registro_feedback']);
                if (shAgenteHis) {
                    let d = obterValoresSeguros(shAgenteHis);
                    for (let i = 1; i < d.length; i++) {
                        let idColab = String(d[i][0] || '').trim();
                        let nomeColab = String(d[i][1] || '').trim();
                        if (!idColab && !nomeColab) continue;

                        let idL = normalizarFilialId(d[i][2]);
                        let statusDesl = String(d[i][17] || '').toUpperCase();
                        if (statusDesl.includes('DESLIGADO') || (idColab && dSet.has(idColab))) continue;
                        
                        if (idL && masterData[idL]) {
                            masterData[idL].agente.elegiveis++;
                            let dataAdm = extrairDataBR(d[i][5]); 
                            let diasC = parseInt(d[i][6]) || 0; 
                            let f15 = String(d[i][8]).trim() === "1"; 
                            let f30 = String(d[i][9]).trim() === "1"; 
                            let f45 = String(d[i][10]).trim() === "1"; 

                            let missing = [];
                            if(diasC >= 15 && !f15) missing.push('15d');
                            if(diasC >= 30 && !f30) missing.push('30d');
                            if(diasC >= 45 && !f45) missing.push('45d');

                            if (missing.length > 0) { masterData[idL].pendencias.agenteFeedbacks.push({ nome: nomeColab, dataAdm: dataAdm, diasTempo: diasC, faltantes: missing.join(', ') }); } 
                            else { masterData[idL].agente.concluidos++; }
                        }
                    }
                }
            } catch(e) {}
            console.timeEnd("Leitura_DB_AGENTE");
        } catch (e) {
            avisos.push("Base Agente inacessível: " + e.message);
            console.warn("Aviso DB_AGENTE:", e.message);
        }

        try {
            console.time("Leitura_DB_SOC");
            const ssSoc = SpreadsheetApp.openById(DB_SOC);
            
            const procSoc = (n, tipo) => {
                let sh = getSheetFast(ssSoc, n);
                if(sh) {
                    let d = obterValoresSeguros(sh);
                    for(let i=1; i<d.length; i++) {
                        let idL = normalizarFilialId(d[i][2]);
                        let statusSit = String(d[i][6] || '').toUpperCase();
                        if(idL && masterData[idL] && !statusSit.includes('CANCELADO')) {
                            let m = extrairMesAno(d[i][0]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                            masterData[idL].pendencias.socialTodos.push({ mes: m, data: extrairDataBR(d[i][0]), fluxo: tipo, colaborador: String(d[i][3]||'').trim(), motivo: String(d[i][8]||'').trim(), situacao: statusSit, retomada: extrairDataBR(d[i][11]), parecer: String(d[i][10]||'').trim() });
                            injetarAcaoVirtual(idL, d[i][0], 'Atendimento Social', 'Acompanhamento Psicológico/Social', 1, 'SOCIAL');
                        }
                    }
                }
            };
            procSoc(['base_registro'], 'EQUIPE SS');
            procSoc(['base_historico'], 'EQUIPE SS');
            procSoc(['base_internos'], 'GP');
            console.timeEnd("Leitura_DB_SOC");
        } catch(e) {
            avisos.push("Base Social inacessível: " + e.message);
            console.warn("Aviso DB_SOC:", e.message);
        }

        let pMeses = Array.from(filtrosGlobais.meses).filter(Boolean).filter(m => { return m >= '2022-01' && m <= hojeIsoMes; }).sort().reverse();
        let pDir = Array.from(filtrosGlobais.diretorias).filter(Boolean).sort();
        let pReg = Array.from(filtrosGlobais.regionais).filter(Boolean).sort();
        let pFil = Array.from(filtrosGlobais.filiais).filter(Boolean).sort((a,b) => parseInt(a) - parseInt(b));
        let pCoord = Array.from(filtrosGlobais.coordenadores).filter(Boolean).sort();

        console.timeEnd("CargaMestre_Total");
        console.log("Carga mestre gerada com sucesso! Lojas:", Object.keys(masterData).length, "Ações mural:", acoesMuralGlobais.length);

        return JSON.stringify({ 
            sucesso: true, acesso: acesso, 
            filtrosGlobais: { meses: pMeses, coordenadores: pCoord, diretorias: pDir, regionais: pReg, filiais: pFil }, 
            coordFotos: coordFotos, dadosUnificados: masterData, acoesGlobais: acoesMuralGlobais, avisosSistema: avisos 
        });

    } catch (err) {
        console.error("Erro Crítico em obterCargaMestreDashboard:", err);
        return JSON.stringify({ sucesso: false, erro: "Erro Crítico: " + err.message });
    }
}
