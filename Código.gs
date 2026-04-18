/**
 * ============================================================================
 * BACKEND: RINCON OPS (Google Apps Script)
 * Arquitetura Modular & Motor Anti-Conflito O(1) + Sheets API V4
 * ============================================================================
 */

// 1. INJEÇÃO DE DEPENDÊNCIAS (MODULARIZAÇÃO)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 2. ENTRY POINT (RENDERIZAÇÃO DO APP)
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Rincon Ops - Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ----------------------------------------------------------------------------
// 3. CACHE E BANCO DE DADOS (GATEKEEPER DE ESTRUTURA)
// ----------------------------------------------------------------------------
let _dbCacheId = null;

// 🔥 O GATEKEEPER: Garante que a estrutura física existe antes da API agir
function buildDatabaseStructure() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('RINCON_OPS_DB');
  let ss = null;
  let isNewBuild = false;

  // 1. Tenta abrir a planilha existente e valida se a aba principal está lá
  if (dbId) {
    try { 
      ss = SpreadsheetApp.openById(dbId); 
      if (!ss.getSheetByName("LOCAIS_EVENTOS")) {
        ss = null; // Arquivo corrompido ou na lixeira
      }
    } 
    catch(e) { ss = null; }
  }
  
  // 2. Constrói do zero se não existir ou estiver corrompido
  if (!ss) {
    isNewBuild = true;
    ss = SpreadsheetApp.create("DB_Rincon_Ops");
    dbId = ss.getId();
    props.setProperty('RINCON_OPS_DB', dbId);
    
    // Constrói todas as abas e cabeçalhos OBRIGATÓRIOS
    setupSheet(ss, "LOCAIS_EVENTOS", ["ID", "Nome", "Tipo", "Data_Inicio", "Hora_Inicio", "Data_Fim", "Hora_Fim", "Status"]);
    setupSheet(ss, "FUNCIONARIOS", ["ID", "Nome", "Telefone", "Status"]);
    setupSheet(ss, "FUNCOES", ["ID", "Nome"]);
    setupSheet(ss, "ESCALAS", ["ID_Escala", "ID_Funcionario", "ID_LocalEvento", "Data", "Horario_Entrada", "Horario_Saida", "Status", "Funcao"]);
    
    // Limpa aba inútil padrão
    const page1 = ss.getSheetByName("Página1");
    if (page1 && ss.getSheets().length > 1) {
      ss.deleteSheet(page1);
    }

    // Trava o servidor para forçar a gravação
    SpreadsheetApp.flush(); 
    
    // PAUSA DE 2 SEGUNDOS: Necessário para a replicação nos servidores do Google Drive
    // Evita o erro "Unable to parse range" na Sheets API
    Utilities.sleep(2000); 
  }
  
  _dbCacheId = dbId;
  return dbId;
}

function setupSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("#f59e0b");
    sheet.setFrozenRows(1);
  }
}

// Retorna apenas a ID (Otimizado)
function getDBId() {
  if (_dbCacheId) return _dbCacheId;
  return buildDatabaseStructure(); 
}

// Retorna o objeto (Usado nas funções antigas)
function getSheet(name) { 
  return SpreadsheetApp.openById(getDBId()).getSheetByName(name); 
}

function sheetToJSON(data) {
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]]))).filter(r => r[headers[0]] !== "");
}

function parseDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  const [y, m, d] = dateStr.split('-');
  const [hr, min] = timeStr.split(':');
  return new Date(y, m - 1, d, hr, min);
}

// ----------------------------------------------------------------------------
// 4. LÓGICA DE NEGÓCIO E APIs (OTIMIZADO COM SHEETS API V4)
// ----------------------------------------------------------------------------

function getDashboardData() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_full_data';
    const cached = cache.get(cacheKey);
    
    if (cached) return JSON.parse(cached);

    // 🔥 Invoca o Gatekeeper ANTES de qualquer coisa
    const ssId = getDBId();
    
    // BATCH GET VIA SHEETS API
    const response = Sheets.Spreadsheets.Values.batchGet(ssId, {
      ranges: ['LOCAIS_EVENTOS!A:H', 'FUNCIONARIOS!A:D', 'ESCALAS!A:H', 'FUNCOES!A:B']
    });

    const locaisRaw = sheetToJSON(response.valueRanges[0].values || [["ID"]]);
    const funcionarios = sheetToJSON(response.valueRanges[1].values || [["ID"]]);
    const escalasRaw = sheetToJSON(response.valueRanges[2].values || [["ID_Escala"]]);
    const funcoesRaw = sheetToJSON(response.valueRanges[3].values || [["ID"]]);
    
    let funcoes = funcoesRaw.length > 0 ? funcoesRaw : [
      {ID: '1', Nome: 'Vigilante'}, {ID: '2', Nome: 'Segurança Tático'}, {ID: '3', Nome: 'Chefe de Equipa'}
    ];

    const now = new Date();
    const locais = locaisRaw.map(local => {
      let expired = false;
      if (local.Status !== 'Cancelado' && local.Tipo === 'Evento' && local.Data_Fim && local.Hora_Fim) {
        const endDate = parseDateTime(local.Data_Fim, local.Hora_Fim);
        if (endDate && endDate < now) { expired = true; local.Status = 'Expirado'; }
      }
      return { ...local, expired };
    });

    const escalas = escalasRaw.map(e => ({...e, Status: e.Status || 'Confirmado', Funcao: e.Funcao || 'Vigilante'}));
    const result = { success: true, data: { locais, funcionarios, escalas, funcoes } };
    
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(e) {}
    return result;
  } catch (e) { return { success: false, error: "Erro na API: " + e.message }; }
}

function invalidateDashboardCache() {
  try { CacheService.getScriptCache().remove('dashboard_full_data'); } catch(e) {}
}

function buildEscalasIndex(escalas) {
  const index = {};
  escalas.forEach(esc => {
    if (!index[esc.ID_Funcionario]) index[esc.ID_Funcionario] = [];
    index[esc.ID_Funcionario].push(esc);
  });
  return index;
}

function validateConflictMemoryOptimized(idFuncionario, dateStr, startStr, endStr, ignoreId, escalasIndex) {
  const newStart = parseDateTime(dateStr, startStr);
  let newEnd = parseDateTime(dateStr, endStr);
  if (!newStart || !newEnd) throw new Error("Data ou horário inválido.");
  if (newEnd <= newStart) newEnd.setDate(newEnd.getDate() + 1);

  const funcionarioEscalas = escalasIndex[idFuncionario] || [];
  for (let esc of funcionarioEscalas) {
    if (esc.Status === 'Cancelado') continue;
    if (ignoreId && esc.ID_Escala === ignoreId) continue;

    let escStart = parseDateTime(esc.Data, esc.Horario_Entrada);
    let escEnd = parseDateTime(esc.Data, esc.Horario_Saida);
    if (escEnd <= escStart) escEnd.setDate(escEnd.getDate() + 1);

    if (newStart < escEnd && newEnd > escStart) {
      throw new Error(`O agente selecionado já possui uma escala sobreposta no dia ${dateStr}.`);
    }
  }
}

// Retém uso do nativo para edições isoladas (mais simples de manter na linha exata)
function updateOrAddRow(sheetName, idColIndex, id, newData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColIndex] === id) { rowIndex = i + 1; break; }
  }
  if (rowIndex > -1) sheet.getRange(rowIndex, 1, 1, newData.length).setValues([newData]); 
  else sheet.appendRow(newData); 
}

function saveFuncao(nome) {
  try {
    const id = Utilities.getUuid();
    getSheet("FUNCOES").appendRow([id, nome]);
    invalidateDashboardCache();
    return { success: true, data: { ID: id, Nome: nome } };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveFuncionario(payload) {
  try {
    const id = payload.ID || Utilities.getUuid();
    updateOrAddRow("FUNCIONARIOS", 0, id, [id, payload.Nome, payload.Telefone, payload.Status || 'Ativo']);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveLocal(payload) {
  try {
    const id = payload.ID || Utilities.getUuid();
    updateOrAddRow("LOCAIS_EVENTOS", 0, id, [id, payload.Nome, payload.Tipo, payload.Data_Inicio||'', payload.Hora_Inicio||'', payload.Data_Fim||'', payload.Hora_Fim||'', payload.Status || 'Ativo']);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveEscalaBatch(payloads) {
  try {
    // 🔥 Garante a estrutura e pega o ID
    const ssId = getDBId();
    
    // API para leitura rápida do estado atual
    const response = Sheets.Spreadsheets.Values.get(ssId, 'ESCALAS!A:H');
    const escalasAtuais = sheetToJSON(response.values || [["ID_Escala"]]);
    const escalasIndex = buildEscalasIndex(escalasAtuais);

    for (let payload of payloads) {
      validateConflictMemoryOptimized(payload.ID_Funcionario, payload.Data, payload.Horario_Entrada, payload.Horario_Saida, payload.ID_Escala, escalasIndex);
    }

    // BATCH UPDATE VIA SHEETS API
    let newRows = [];
    
    for (let payload of payloads) {
      if (!payload.ID_Escala) {
        newRows.push([
          Utilities.getUuid(), payload.ID_Funcionario, payload.ID_LocalEvento, 
          payload.Data, payload.Horario_Entrada, payload.Horario_Saida, 
          payload.Status || 'Confirmado', payload.Funcao || 'Vigilante'
        ]);
      } else {
        updateOrAddRow("ESCALAS", 0, payload.ID_Escala, [
          payload.ID_Escala, payload.ID_Funcionario, payload.ID_LocalEvento, 
          payload.Data, payload.Horario_Entrada, payload.Horario_Saida, 
          payload.Status || 'Confirmado', payload.Funcao || 'Vigilante'
        ]);
      }
    }

    // Insere registros em massa instantaneamente
    if (newRows.length > 0) {
      Sheets.Spreadsheets.Values.append(
        { values: newRows }, 
        ssId, 
        'ESCALAS!A:H', 
        { valueInputOption: 'USER_ENTERED' }
      );
    }

    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function deleteRow(sheetName, id) {
  try {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { rowIndex = i + 1; break; }
    }
    if (rowIndex > -1) sheet.deleteRow(rowIndex);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}
