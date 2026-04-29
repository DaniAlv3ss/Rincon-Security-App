/**
 * ============================================================================
 * BACKEND: RINCON SECURITY (Google Apps Script)
 * Arquitetura Modular & Motor Anti-Conflito O(1) + Sheets API V4
 * ============================================================================
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Rincon Security: Gestão de Escala')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

let _dbCacheId = null;

function buildDatabaseStructure() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('RINCON_OPS_DB');
  let ss = null;

  if (dbId) {
    try { 
      ss = SpreadsheetApp.openById(dbId); 
      if (!ss.getSheetByName("LOCAIS_EVENTOS")) ss = null;
    } catch(e) { ss = null; }
  }
  
  if (!ss) {
    ss = SpreadsheetApp.create("DB_Rincon_Ops");
    dbId = ss.getId();
    props.setProperty('RINCON_OPS_DB', dbId);
    
    setupSheet(ss, "LOCAIS_EVENTOS", ["ID", "Nome", "Tipo", "Data_Inicio", "Hora_Inicio", "Data_Fim", "Hora_Fim", "Status", "EnderecoLink", "CEP", "Cidade", "Estado", "TipoFixoDuracao"]);
    setupSheet(ss, "FUNCIONARIOS", ["ID", "Nome", "Telefone", "Status"]);
    setupSheet(ss, "FUNCOES", ["ID", "Nome", "Valor", "Descricao"]);
    setupSheet(ss, "ESCALAS", ["ID_Escala", "ID_Funcionario", "ID_LocalEvento", "Data", "Horario_Entrada", "Horario_Saida", "Status", "Funcao", "Data_Fim", "DataHoraPagamento", "Uniforme", "Escopo", "Contato"]);
    
    const page1 = ss.getSheetByName("Página1");
    if (page1 && ss.getSheets().length > 1) ss.deleteSheet(page1);

    SpreadsheetApp.flush(); 
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

function getDBId() {
  if (_dbCacheId) return _dbCacheId;
  return buildDatabaseStructure(); 
}

function getSheet(name) { 
  return SpreadsheetApp.openById(getDBId()).getSheetByName(name); 
}

function sheetToJSON(data) {
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] !== undefined ? row[i] : ""]))).filter(r => r[headers[0]] !== "");
}

function parseDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  const [y, m, d] = dateStr.split('-');
  const [hr, min] = timeStr.split(':');
  return new Date(y, m - 1, d, hr, min);
}

function getDashboardData() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_full_data';
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    const ssId = getDBId();
    
    const response = Sheets.Spreadsheets.Values.batchGet(ssId, {
      ranges: ['LOCAIS_EVENTOS!A:M', 'FUNCIONARIOS!A:D', 'ESCALAS!A:M', 'FUNCOES!A:D']
    });

    const locaisRaw = sheetToJSON(response.valueRanges[0].values || [["ID"]]);
    const funcionarios = sheetToJSON(response.valueRanges[1].values || [["ID"]]);
    const escalasRaw = sheetToJSON(response.valueRanges[2].values || [["ID_Escala"]]);
    const funcoesRaw = sheetToJSON(response.valueRanges[3].values || [["ID"]]);
    
    let funcoes = funcoesRaw;

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

function validateConflictMemoryOptimized(idFuncionario, dateStr, startStr, dateEndStr, endStr, ignoreId, escalasIndex) {
  const newStart = parseDateTime(dateStr, startStr);
  let newEnd = dateEndStr ? parseDateTime(dateEndStr, endStr) : parseDateTime(dateStr, endStr);
  
  if (!newStart || !newEnd) throw new Error("Data ou horário inválido.");
  if (!dateEndStr && newEnd <= newStart) newEnd.setDate(newEnd.getDate() + 1);

  const funcionarioEscalas = escalasIndex[idFuncionario] || [];
  for (let esc of funcionarioEscalas) {
    if (esc.Status === 'Cancelado') continue;
    if (ignoreId && esc.ID_Escala === ignoreId) continue;

    let escStart = parseDateTime(esc.Data, esc.Horario_Entrada);
    let escEnd = esc.Data_Fim ? parseDateTime(esc.Data_Fim, esc.Horario_Saida) : parseDateTime(esc.Data, esc.Horario_Saida);
    if (!esc.Data_Fim && escEnd <= escStart) escEnd.setDate(escEnd.getDate() + 1);

    if (newStart < escEnd && newEnd > escStart) {
      throw new Error(`Conflito: Agente já possui escala sobreposta.`);
    }
  }
}

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

function saveFuncao(payload) {
  try {
    const id = payload.ID || Utilities.getUuid();
    updateOrAddRow("FUNCOES", 0, id, [id, payload.Nome, payload.Valor || '', payload.Descricao || '']);
    invalidateDashboardCache();
    return { success: true, data: { ID: id, Nome: payload.Nome, Valor: payload.Valor, Descricao: payload.Descricao } };
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
    const rowData = [
      id, payload.Nome, payload.Tipo, payload.Data_Inicio||'', payload.Hora_Inicio||'', payload.Data_Fim||'', payload.Hora_Fim||'', payload.Status || 'Ativo',
      payload.EnderecoLink || '', payload.CEP || '', payload.Cidade || '', payload.Estado || '', payload.TipoFixoDuracao || ''
    ];
    updateOrAddRow("LOCAIS_EVENTOS", 0, id, rowData);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveEscalaBatch(payloads) {
  try {
    const ssId = getDBId();
    const response = Sheets.Spreadsheets.Values.get(ssId, 'ESCALAS!A:M');
    const escalasAtuais = sheetToJSON(response.values || [["ID_Escala"]]);
    const escalasIndex = buildEscalasIndex(escalasAtuais);

    for (let payload of payloads) {
      validateConflictMemoryOptimized(payload.ID_Funcionario, payload.Data, payload.Horario_Entrada, payload.Data_Fim, payload.Horario_Saida, payload.ID_Escala, escalasIndex);
    }

    let newRows = [];
    
    for (let payload of payloads) {
      const rowData = [
        payload.ID_Escala || Utilities.getUuid(), payload.ID_Funcionario, payload.ID_LocalEvento, 
        payload.Data, payload.Horario_Entrada, payload.Horario_Saida, 
        payload.Status || 'Confirmado', payload.Funcao || 'Vigilante',
        payload.Data_Fim || '', payload.DataHoraPagamento || '', payload.Uniforme || '', payload.Escopo || '', payload.Contato || ''
      ];

      if (!payload.ID_Escala) {
        newRows.push(rowData);
      } else {
        updateOrAddRow("ESCALAS", 0, payload.ID_Escala, rowData);
      }
    }

    if (newRows.length > 0) {
      Sheets.Spreadsheets.Values.append(
        { values: newRows }, ssId, 'ESCALAS!A:M', { valueInputOption: 'USER_ENTERED' }
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
