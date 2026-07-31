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

function dataParaStringISO(dataVal, fallbackIsoData = null) {
    if (!dataVal || String(dataVal).trim() === "") return fallbackIsoData;
    try {
        if (dataVal instanceof Date) {
            return Utilities.formatDate(dataVal, "America/Sao_Paulo", "yyyy-MM-dd");
        }

        let str = String(dataVal).trim();
        
        if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
            return str.substring(0, 10);
        }

        let cleanStr = str.replace(/Horário Padrão de Brasília/gi, '').replace(/GMT.*/, '').trim();
        
        let matchBR = cleanStr.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
        if (matchBR) {
            return `${matchBR[3]}-${matchBR[2]}-${matchBR[1]}`;
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

function getSheetSafe(ss, names) {
    let sheets = ss.getSheets();
    let targetNames = names.map(normalizeName);
    for (let sh of sheets) {
        let currentName = normalizeName(sh.getName());
        if (targetNames.includes(currentName)) return sh;
    }
    return null;
}

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

function obterCargaMestreDashboard() {
    let avisos = [];
    try {
        const emailLogado = Session.getActiveUser().getEmail().toLowerCase().trim();
        const isAdmin = (emailLogado === 'fabiane.satie@magazineluiza.com.br');
        const canViewRoi = (emailLogado === 'fabiane.satie@magazineluiza.com.br' || emailLogado === 'tarcisio.maniglia@magazineluiza.com.br');
        
        let regionaisPermitidas = new Set();
        let role = 'COORDENADOR';
        let acesso = { email: emailLogado, isAdmin: isAdmin, canViewRoi: canViewRoi, role: role };
        
        let masterData = {}; 
        let acoesMuralGlobais = [];
        let visitTracker = new Set(); // Engine de Deduplicação
        
        let filtrosGlobais = { meses: new Set(), diretorias: new Set(), regionais: new Set(), filiais: new Set(), coordenadores: new Set() };
        let coordFotos = {};
        let emailToName = {};

        let hojeIsoObj = new Date();
        let hojeIsoDia = Utilities.formatDate(hojeIsoObj, "America/Sao_Paulo", "yyyy-MM-dd");
        let hojeIsoMes = Utilities.formatDate(hojeIsoObj, "America/Sao_Paulo", "yyyy-MM");

        let ssMaster;
        try { ssMaster = SpreadsheetApp.openById(DB_MASTER); } 
        catch(e) { return JSON.stringify({ sucesso: false, erro: `Falha ao acessar GP360. Permissão negada.` }); }
        
        let injetarAcaoVirtual = (idFilial, dataRaw, motivo, obs, peso, origem = 'VIRTUAL', autorEspecifico = null) => {
            if(!idFilial || !masterData[idFilial]) return;
            let m = extrairMesAno(dataRaw);
            let autor = autorEspecifico || masterData[idFilial].coordenadores[0] || 'Sistema';
            acoesMuralGlobais.push({
                idFilial: idFilial, destinoBruto: idFilial, mes: m, dataStr: extrairDataBR(dataRaw), dataIso: dataParaStringISO(dataRaw),
                autor: autor, motivo: motivo, obs: obs,
                gastos: 0, km: 0, moedas: peso, pessoas: 1, tempo: 1,
                isVisita: false, isVisitaUnica: false, isSocial: motivo.toUpperCase().includes('SOCIAL'),
                origemVirtual: origem 
            });
        };

        const shUsu = ssMaster.getSheetByName('DADOS_USUARIOS');
        let coordMap = {}; 

        if (shUsu) {
            const dadosUsu = shUsu.getDataRange().getValues();
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
                    else if (isAdmin || nivelAcesso.includes('MASTER')) acesso.role = 'ADMIN';
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
        
        if (!isAdmin && regionaisPermitidas.size === 0 && acesso.role !== 'DIRETORRH' && acesso.role !== 'GERENTEGP') {
            return JSON.stringify({ sucesso: false, erro: `Conta sem regionais autorizadas.` });
        }

        const shLojas = ssMaster.getSheetByName('DADOS_LOJAS');
        if (!shLojas) return JSON.stringify({ sucesso: false, erro: "Aba DADOS_LOJAS não encontrada." });
        
        const dadosLojas = shLojas.getDataRange().getValues();
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
            const dadosLanc = shLanc.getDataRange().getValues();

            for(let i = 1; i < dadosLanc.length; i++) {
                let dataAcao = dadosLanc[i][12]; 
                if (!dataAcao) continue;

                let destinoFilialBruto = String(dadosLanc[i][3] || '').trim(); 
                let idL = normalizarFilialId(destinoFilialBruto); 

                let motivoVal = String(dadosLanc[i][4] || '').trim(); // Motivo_Meta
                let tipoRoteiro = String(dadosLanc[i][17] || ''); 
                
                let isVisita = tipoRoteiro.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().includes("visitainloco") || tipoRoteiro.toUpperCase().includes("PRESENCIAL");
                let roteiroLimpo = tipoRoteiro.toUpperCase();
                let motivoUpper = motivoVal.toUpperCase();

                let isSocial = motivoUpper.includes('SOCIAL') || motivoUpper.includes('ASSIST') || roteiroLimpo.includes('SOCIAL') || roteiroLimpo.includes('ASSIST');
                
                // NOVO: Validação de Lançamento Estritamente Financeiro (Avulso)
                let isAvulso = destinoFilialBruto.toUpperCase().startsWith('AVULSO') || motivoUpper.includes('AVULSO');

                let mesStr = extrairMesAno(dataAcao);
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
                    let acaoObj = {
                        idFilial: idL, destinoBruto: destinoFilialBruto, mes: mesStr, dataStr: extrairDataBR(dataAcao), dataIso: isoAcao, autor: nomeAutorCorreto,
                        motivo: motivoVal, obs: String(dadosLanc[i][10] || '').trim(),
                        gastos: parseFloat(String(dadosLanc[i][21] || '').replace(',','.')) || 0, 
                        km: parseFloat(String(dadosLanc[i][15] || '').replace(',','.')) || 0, 
                        
                        // Zera o peso se for apenas deslocamento avulso
                        moedas: isAvulso ? 0 : (parseFloat(String(dadosLanc[i][9] || '').replace(/\./g, '').replace(',','.')) || 0), 
                        pessoas: isAvulso ? 0 : (Number(dadosLanc[i][19]) || 0), 
                        tempo: isAvulso ? 0 : (Number(dadosLanc[i][20]) || 0),
                        
                        isVisita: isVisita && !isAvulso, isVisitaUnica: isVisitaUnica, isSocial: isSocial && !isAvulso, roteiro: tipoRoteiro, origemVirtual: 'FISICA',
                        isAvulsoFinanceiro: isAvulso
                    };

                    acoesMuralGlobais.push(acaoObj);

                    if (idL && masterData[idL]) {
                        masterData[idL].historicoGP.push(acaoObj);
                    }
                }
            }
        }

        try {
            const ssDash = SpreadsheetApp.openById(DB_DASH);
            
            let mapDash = (shNames, parser) => { let sh = getSheetSafe(ssDash, shNames); if(sh) { let d = sh.getDataRange().getValues(); for(let i=1; i<d.length; i++) parser(d[i]); } };
            
            // Sempre carregar tabelas de Resumo Filial (Painel Geral)
            mapDash(['HISTORICO_VENDA_ANO'], r => { let idL = normalizarFilialId(r[0]); if(idL && masterData[idL]) { let atingAcumulado = parseFloat(String(r[7]).replace('%','').replace(',','.')); if(!isNaN(atingAcumulado)) masterData[idL].indicadores.venda = (atingAcumulado < 10) ? atingAcumulado : (atingAcumulado/100); } });
            mapDash(['HISTORICO_TXDESL', 'HISTORICO_INDICADORES_GP'], r => { let idL = normalizarFilialId(r[0]); if(idL && masterData[idL]) { let tx = parseFloat(String(r[9]).replace('%','').replace(',','.')); if(!isNaN(tx)) masterData[idL].indicadores.txSaida = (tx < 10) ? tx : (tx/100); } });
            mapDash(['HISTORICO_QUADRO'], d => { let idL = normalizarFilialId(d[0]); let cargo = String(d[7] || '').trim(); let cargoUpper = cargo.toUpperCase(); if(idL && masterData[idL] && !cargoUpper.includes('INTERMITENTE') && !cargoUpper.includes('MONTAGEM') && !cargoUpper.includes('MONTADOR')) { let v = parseFloat(d[12]); if(!isNaN(v) && v > 0) { masterData[idL].pendencias.vagasInfo.push({ cargo: cargo, qtd: v }); } } });
            
            let mapBHAux = {};
            mapDash(['HISTORICO_BH_ACUMULADO'], d => {
                let idL = normalizarFilialId(d[1]); let dtIso = dataParaStringISO(d[0], hojeIsoDia);
                if(idL && masterData[idL] && dtIso <= hojeIsoDia) { 
                    let hrs = parseFloat(String(d[9]).replace(',','.'));
                    if(!isNaN(hrs)) {
                        if(!mapBHAux[idL]) mapBHAux[idL] = { maxDate: dtIso, total: 0, criticos: [] };
                        if(dtIso > mapBHAux[idL].maxDate) { mapBHAux[idL].maxDate = dtIso; mapBHAux[idL].total = 0; mapBHAux[idL].criticos = []; }
                        if(dtIso === mapBHAux[idL].maxDate) { mapBHAux[idL].total += hrs; if(hrs >= 5) mapBHAux[idL].criticos.push({ nome: String(d[6]).trim(), id: String(d[5]).trim(), horas: hrs }); }
                    }
                }
            });
            Object.keys(mapBHAux).forEach(idL => { masterData[idL].indicadores.bancoHoras = mapBHAux[idL].total; masterData[idL].pendencias.bancoHoras = mapBHAux[idL].criticos; });

            let mapNPS = {}; 
            mapDash(['HISTORICO_NPS'], d => { let idL = normalizarFilialId(d[1]); let dtIso = dataParaStringISO(d[0], hojeIsoDia); if(idL && masterData[idL] && dtIso <= hojeIsoDia) { let npsVal = parseFloat(String(d[6]).replace(',','.')); if(!isNaN(npsVal)) { if(!mapNPS[idL] || dtIso > mapNPS[idL].date) { mapNPS[idL] = { date: dtIso, val: npsVal }; } } } });
            Object.keys(mapNPS).forEach(idL => { masterData[idL].indicadores.nps = mapNPS[idL].val; });

            let registrosNPSV = {}; 
            mapDash(['HISTORICO_NPS_VENDEDORES'], d => { let nome = String(d[1]).trim(); let idL = normalizarFilialId(d[2]); let dtIso = dataParaStringISO(d[0], hojeIsoDia); let npsV = parseFloat(String(d[7]).replace(',','.')); if(idL && nome && masterData[idL] && !isNaN(npsV) && dtIso <= hojeIsoDia) { if(!registrosNPSV[idL]) registrosNPSV[idL] = { maxDate: dtIso, records: [] }; if(dtIso > registrosNPSV[idL].maxDate) { registrosNPSV[idL].maxDate = dtIso; registrosNPSV[idL].records = []; } if(dtIso === registrosNPSV[idL].maxDate) { registrosNPSV[idL].records.push({ nome: nome, nota: npsV }); } } });
            Object.keys(registrosNPSV).forEach(idL => { registrosNPSV[idL].records.forEach(r => { if(r.nota < 70) masterData[idL].pendencias.npsVendedores.push({ nome: r.nome, nota: r.nota }); }); });

        } catch(e) { avisos.push("Base Dash inacessível: " + e.message); }

        try {
            const ssApur = SpreadsheetApp.openById(DB_APUR);
            let mapApur = (shNames, parser) => { let sh = getSheetSafe(ssApur, shNames); if(sh) { let d = sh.getDataRange().getValues(); for(let i=1; i<d.length; i++) { try { parser(d[i]); } catch(e){} } } };

            mapApur(['HISTORICO_APURACAO'], r => {
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
            mapApur(['Rascunhos_Apuracoes'], r => {
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
            mapApur(['Intervencoes_Feedback'], r => {
                let idL = normalizarFilialId(r[1]); let concl = String(r[3] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    masterData[idL].pendencias.intervencoes.push({ data: extrairDataBR(r[0]), linkDoc: String(r[2]||''), status: String(r[3]||''), detalhe: String(r[4]||''), nHumor: String(r[5]||'-'), nData: extrairDataBR(r[6]) });
                    injetarAcaoVirtual(idL, r[0], 'Acompanhamento Feedbacks', 'Intervenção PAI', 1, 'INTERVENCAO');
                }
            });
            mapApur(['Historico_Desligamentos'], r => {
                let idL = normalizarFilialId(r[2]); let concl = String(r[11] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    let m = extrairMesAno(r[1]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                    masterData[idL].pendencias.desligamentos.push({ mes: m, data: extrairDataBR(r[1]), colab: String(r[3]||''), reg: String(r[8]||''), dir: String(r[11]||''), dossieLink: String(r[6]||'') });
                    injetarAcaoVirtual(idL, r[1], 'Relatório de Desligamento', 'Dossiê de Desligamento', 1, 'DESLIGAMENTO');
                }
            });
            mapApur(['Feedbacks_Gerentes'], r => {
                let idL = normalizarFilialId(r[1]); let concl = String(r[5] || '').toUpperCase();
                if(idL && masterData[idL] && !concl.includes('CANCELADO') && !concl.includes('LANÇADO INCORRETAMENTE')) {
                    let m = extrairMesAno(r[2]); if(m!=="0000-00") filtrosGlobais.meses.add(m);
                    masterData[idL].pendencias.feedbacksSLA.push({ mes: m, data: extrairDataBR(r[2]), gerente: String(r[9]||'').trim(), status: String(r[5]||''), dataRetorno: extrairDataBR(r[6]), evidenciaLink: String(r[8]||'') });
                }
            });
        } catch(e) { }

        try {
            const ssAgente = SpreadsheetApp.openById(DB_AGENTE);
            let dSet = new Set();
            try {
                let shDesl = getSheetSafe(ssAgente, ['atualizacao_desligamento', 'desligamentos']);
                if (shDesl) { let d = shDesl.getDataRange().getValues(); for(let i=1; i<d.length; i++) if(d[i][0]) dSet.add(String(d[i][0]).trim()); }
            } catch(e){}

            try {
                let shAgenteHis = getSheetSafe(ssAgente, ['REGISTRO_FEEDBACK']);
                if (shAgenteHis) {
                    let d = shAgenteHis.getDataRange().getValues();
                    for (let i = 1; i < d.length; i++) {
                        let idColab = String(d[i][0] || '').trim();
                        let idL = normalizarFilialId(d[i][2]);
                        
                        let statusDesl = String(d[i][17] || '').toUpperCase();
                        if (statusDesl.includes('DESLIGADO') || (idColab && dSet.has(idColab))) continue;
                        
                        if (idL && masterData[idL]) {
                            masterData[idL].agente.elegiveis++;
                            let nomeColab = String(d[i][1] || '').trim(); 
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
        } catch (e) { }

        try {
            const ssSoc = SpreadsheetApp.openById(DB_SOC);
            const procSoc = (n, tipo) => {
                let sh = getSheetSafe(ssSoc, n);
                if(sh) {
                    let d = sh.getDataRange().getValues();
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
        } catch(e) {}

        let pMeses = Array.from(filtrosGlobais.meses).filter(Boolean).filter(m => { return m >= '2022-01' && m <= hojeIsoMes; }).sort().reverse();
        let pDir = Array.from(filtrosGlobais.diretorias).filter(Boolean).sort();
        let pReg = Array.from(filtrosGlobais.regionais).filter(Boolean).sort();
        let pFil = Array.from(filtrosGlobais.filiais).filter(Boolean).sort((a,b) => parseInt(a) - parseInt(b));
        let pCoord = Array.from(filtrosGlobais.coordenadores).filter(Boolean).sort();

        return JSON.stringify({ 
            sucesso: true, acesso: acesso, 
            filtrosGlobais: { meses: pMeses, coordenadores: pCoord, diretorias: pDir, regionais: pReg, filiais: pFil }, 
            coordFotos: coordFotos, dadosUnificados: masterData, acoesGlobais: acoesMuralGlobais, avisosSistema: avisos 
        });

    } catch (err) { return JSON.stringify({ sucesso: false, erro: "Erro Crítico: " + err.message }); }
}
