/**
 * ECOSSISTEMA GP360 - PORTAL GP 360
 * Arquivo: Service_Write.gs
 * Subpasta Monorepo: src/portal-gp360/
 * Módulo de Gravação Segura no BDD Master com suporte a LockService
 */

/**
 * Normalização global de Filial ID
 * @param {any} val - ID da filial
 * @return {string} ID higienizado
 */
function normalizarFilialId(val) {
  if (!val && val !== 0) return "";
  var num = parseInt(String(val).replace(/\D/g, ''), 10);
  if (isNaN(num)) return String(val).trim();
  if (num > 3000) { num -= 3000; }
  return String(num);
}

/**
 * Salva arquivo de imagem/evidência na pasta do Google Drive
 * @param {Object} arquivoObj - Objeto contendo base64, mimeType e nome
 * @return {string} URL pública ou do Drive do arquivo salvo
 */
function salvarEvidenciaNoDrive(arquivoObj) {
  if (!arquivoObj || !arquivoObj.base64) return '';
  try {
    var folder = DriveApp.getFolderById(EVIDENCIAS_FOLDER_ID);
    var bytes = Utilities.base64Decode(arquivoObj.base64.split(',')[1] || arquivoObj.base64);
    var blob = Utilities.newBlob(bytes, arquivoObj.mimeType || 'image/jpeg', arquivoObj.nome || ('Evidencia_' + new Date().getTime() + '.jpg'));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://lh3.googleusercontent.com/d/' + file.getId() + '=w1000';
  } catch (err) {
    Logger.log('Erro ao salvar evidencia no Drive: ' + err.message);
    return '';
  }
}

/**
 * Salva uma nova atividade no BDD Master (DADOS_LANCAMENTOS)
 * @param {Object} dados - Objeto do formulário enviado pelo front-end
 * @return {Object} Resposta de sucesso ou falha
 */
function salvarAtividadeServidor(dados) {
  var lock = LockService.getScriptLock();
  try {
    // Trava de segurança ajustada usando tryLock
    var success = lock.tryLock(15000);
    if (!success) {
      return { sucesso: false, mensagem: 'O sistema está processando outro salvamento. Tente novamente em alguns segundos.' };
    }

    var controle = obterControleAcesso();
    if (!controle.temAcesso) {
      return { sucesso: false, mensagem: 'Acesso negado. Usuário não autorizado.' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!abaLanc) {
      return { sucesso: false, mensagem: 'Aba DADOS_LANCAMENTOS não encontrada no banco de dados.' };
    }

    // Processamento de evidência
    var urlEvidenciaFinal = dados.evidenciaUrl || '';
    if (dados.arquivoEvidencia && dados.arquivoEvidencia.base64) {
      urlEvidenciaFinal = salvarEvidenciaNoDrive(dados.arquivoEvidencia);
    }

    // Identificador único de registro
    var novoId = 'REG_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
    
    // Normalização da Filial
    var numFilial = normalizarFilialId(dados.filial);
    var filialPadronizada = numFilial ? ("0000" + numFilial).slice(-4) : String(dados.filial || '').trim().toUpperCase();

    // Formatação de data
    var dtData = new Date(dados.dataAtividade + 'T12:00:00');
    var dtStrFormatted = ("0" + dtData.getDate()).slice(-2) + '/' + ("0" + (dtData.getMonth() + 1)).slice(-2) + '/' + dtData.getFullYear();

    // Valores financeiros
    var rateKm = parseFloat(dados.rateKm) || 1.20;
    var km = parseFloat(dados.km) || 0;
    var custoKm = km * rateKm;
    var alimentacao = parseFloat(dados.alimentacao) || 0;
    var hospedagem = parseFloat(dados.hospedagem) || 0;
    var aereo = parseFloat(dados.aereo) || 0;
    var pedagio = parseFloat(dados.pedagio) || 0;
    var estacionamento = parseFloat(dados.estacionamento) || 0;
    var custoTotal = custoKm + alimentacao + hospedagem + aereo + pedagio + estacionamento;

    // Status inicial e validação de especialista
    var motivoLower = String(dados.natureza || '').toLowerCase();
    var ehEspecialista = motivoLower.includes('atendimento social') || 
                         motivoLower.includes('apuraç') || 
                         motivoLower.includes('apurac') || 
                         motivoLower.includes('feedback') || 
                         motivoLower.includes('acompanhamento');

    var statusInicial = ehEspecialista ? 'PENDENTE' : 'VALIDADO';

    // Montagem da linha de dados para inserção
    var novaLinha = [
      novoId,                          // Coluna A: ID Registro
      dtStrFormatted,                   // Coluna B: Data Formatada (DD/MM/YYYY)
      controle.email,                   // Coluna C: Autor Email
      filialPadronizada,                // Coluna D: Filial Destino
      dados.natureza || '',             // Coluna E: Motivo / Natureza
      '',                               // Coluna F: Reservado
      '',                               // Coluna G: Reservado
      '',                               // Coluna H: Reservado
      custoTotal,                       // Coluna I: Custo Total
      1,                                // Coluna J: Pontos/Moeda Padrão
      dados.observacao || '',           // Coluna K: Observações
      urlEvidenciaFinal,                // Coluna L: URL Evidência
      dados.dataAtividade,              // Coluna M: Data Raw ISO (YYYY-MM-DD)
      dados.escopo || 'FILIAL',         // Coluna N: Escopo (FILIAL/REGIONAL/DIRETORIA)
      '',                               // Coluna O: Reservado
      km,                               // Coluna P: KM Percorrido
      custoKm,                          // Coluna Q: Custo KM
      dados.tipoRoteiro || 'Presencial',// Coluna R: Tipo Roteiro
      (dados.temasReuniao || []).join(', ') || dados.temaTreinamento || '', // Coluna S: Temas/Treinamento
      parseFloat(dados.pessoasImpactadas) || 0, // Coluna T: Pessoas Impactadas
      parseFloat(dados.tempoGasto) || 0,         // Coluna U: Tempo Gasto
      custoTotal,                       // Coluna V: Total Reembolso
      alimentacao,                      // Coluna W: Alimentação
      hospedagem,                       // Coluna X: Hospedagem
      aereo,                            // Coluna Y: Aéreo
      pedagio,                          // Coluna Z: Pedágio
      estacionamento,                   // Coluna AA: Estacionamento
      statusInicial                     // Coluna AB: Status Validação
    ];

    abaLanc.appendRow(novaLinha);

    // Limpa caches para garantir atualização rápida dos totais
    var cache = CacheService.getScriptCache();
    cache.remove('GP360_MAPA_INDICADORES_SLIM_V5');

    return { sucesso: true, mensagem: 'Atividade salva com sucesso! +1 Moeda na sua jornada.' };

  } catch (err) {
    Logger.log('Erro ao salvar atividade: ' + err.message);
    return { sucesso: false, mensagem: 'Erro interno no servidor ao salvar: ' + err.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Salva lançamento exclusivo de KM Avulso (Viagem à Sede / Eventos)
 * @param {Object} dados - Dados do formulário de KM Avulso
 * @return {Object} Resposta da operação
 */
function salvarKMAvulsoServidor(dados) {
  var lock = LockService.getScriptLock();
  try {
    var success = lock.tryLock(15000);
    if (!success) {
      return { sucesso: false, mensagem: 'Sistema ocupado. Tente novamente em alguns segundos.' };
    }

    var controle = obterControleAcesso();
    if (!controle.temAcesso) {
      return { sucesso: false, mensagem: 'Acesso negado.' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!abaLanc) {
      return { sucesso: false, mensagem: 'Aba DADOS_LANCAMENTOS não encontrada.' };
    }

    var novoId = 'REG_AVULSO_' + new Date().getTime();
    var dtData = new Date(dados.data + 'T12:00:00');
    var dtStrFormatted = ("0" + dtData.getDate()).slice(-2) + '/' + ("0" + (dtData.getMonth() + 1)).slice(-2) + '/' + dtData.getFullYear();

    var rateKm = parseFloat(dados.rate) || 1.20;
    var km = parseFloat(dados.km) || 0;
    var custoKm = km * rateKm;
    var alimentacao = parseFloat(dados.alimentacao) || 0;
    var hospedagem = parseFloat(dados.hospedagem) || 0;
    var aereo = parseFloat(dados.aereo) || 0;
    var pedagio = parseFloat(dados.pedagio) || 0;
    var estacionamento = parseFloat(dados.estacionamento) || 0;
    var custoTotal = custoKm + alimentacao + hospedagem + aereo + pedagio + estacionamento;

    var novaLinha = [
      novoId, dtStrFormatted, controle.email, 'AVULSO - ' + (dados.destino || 'SEDE'),
      'Deslocamento KM Avulso', '', '', '', custoTotal, 0,
      'KM Avulso com destino a: ' + dados.destino, '', dados.data, 'REGIONAL',
      '', km, custoKm, 'Presencial', 'KM Avulso', 1, 1, custoTotal,
      alimentacao, hospedagem, aereo, pedagio, estacionamento, 'VALIDADO'
    ];

    abaLanc.appendRow(novaLinha);
    return { sucesso: true, mensagem: 'Lançamento de KM Avulso registrado com sucesso!' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao salvar KM Avulso: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Exclui um lançamento do BDD pelo ID
 * @param {string} idRegistro - ID único do registro
 * @return {Object} Status da exclusão
 */
function deletarLancamento(idRegistro) {
  var lock = LockService.getScriptLock();
  try {
    var success = lock.tryLock(15000);
    if (!success) return { sucesso: false, mensagem: 'Servidor ocupado.' };

    var controle = obterControleAcesso();
    if (!controle.temAcesso) return { sucesso: false, mensagem: 'Sem permissão.' };

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
    if (!abaLanc) return { sucesso: false, mensagem: 'Tabela não encontrada.' };

    var dados = abaLanc.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === String(idRegistro).trim()) {
        var emailAutor = String(dados[i][2] || '').toLowerCase();
        if (controle.isSuperAdmin || emailAutor === controle.email.toLowerCase()) {
          abaLanc.deleteRow(i + 1);
          return { sucesso: true, mensagem: 'Atividade excluída com sucesso!' };
        } else {
          return { sucesso: false, mensagem: 'Você só pode excluir seus próprios lançamentos.' };
        }
      }
    }
    return { sucesso: false, mensagem: 'Registro não localizado.' };
  } catch (e) {
    return { sucesso: false, mensagem: 'Erro ao excluir: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Retorna as evidências de fotos cadastradas do usuário para a Galeria
 * @param {number|string} mes - Mês filtro
 * @param {number|string} ano - Ano filtro
 * @return {Array} Lista de fotos com links
 */
function getEvidenciasUsuario(mes, ano) {
  var controle = obterControleAcesso();
  if (!controle.temAcesso) return [];

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var abaLanc = ss.getSheetByName('DADOS_LANCAMENTOS');
  if (!abaLanc) return [];

  var mesAlvo = parseInt(mes, 10) || (new Date().getMonth() + 1);
  var anoAlvo = parseInt(ano, 10) || new Date().getFullYear();

  var fotos = [];
  var dados = abaLanc.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    var row = dados[i];
    var urlEvidencia = String(row[11] || '').trim();
    if (!urlEvidencia) continue;

    var dtInfo = extrairMesAnoData(row[12] || row[1]);
    if (dtInfo.mes === mesAlvo && dtInfo.ano === anoAlvo) {
      var emailAutor = String(row[2] || '').toLowerCase();
      if (controle.isSuperAdmin || emailAutor === controle.email.toLowerCase()) {
        fotos.push({
          id: row[0],
          loja: row[3] || 'AVULSO',
          motivo: row[4] || 'Atividade',
          data: ("0" + dtInfo.dia).slice(-2) + '/' + ("0" + dtInfo.mes).slice(-2) + '/' + dtInfo.ano,
          url: urlEvidencia,
          thumbUrl: urlEvidencia
        });
      }
    }
  }

  return fotos.reverse();
}
