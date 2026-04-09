/**
 * ============================================================================
 * BACKEND: SAFEGUARD ELITE (Google Apps Script)
 * Arquitetura Relacional, Motor Anti-Conflito & Lotes de Agentes
 * ============================================================================
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SafeGuard Elite - Rincon Edition')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ----------------------------------------------------------------------------
// 1. AUTO-SETUP & CACHE GLOBAL (Otimização Extrema)
// ----------------------------------------------------------------------------
let _dbCache = null; // Cache em memória para a execução atual

function getDB() {
  if (_dbCache) return _dbCache; // Retorna imediatamente se já abriu
  
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('SG_ELITE_DB');
  
  if (dbId) {
    try { 
      _dbCache = SpreadsheetApp.openById(dbId); 
    } catch(e) { 
      dbId = null; // Se o ID for inválido ou a planilha foi apagada
    }
  }
  
  // Se não tem banco, cria um novo
  if (!dbId) {
    _dbCache = SpreadsheetApp.create("DB_SafeGuard_Elite");
    props.setProperty('SG_ELITE_DB', _dbCache.getId());
    
    if (_dbCache.getSheets()[0].getName() === "Página1") {
      _dbCache.deleteSheet(_dbCache.getSheets()[0]);
    }
    
    setupSheet(_dbCache, "LOCAIS_EVENTOS", ["ID", "Nome", "Tipo", "Data_Inicio", "Hora_Inicio", "Data_Fim", "Hora_Fim", "Status"]);
    setupSheet(_dbCache, "FUNCIONARIOS", ["ID", "Nome", "Telefone", "Status"]);
    setupSheet(_dbCache, "FUNCOES", ["ID", "Nome"]);
    setupSheet(_dbCache, "ESCALAS", ["ID_Escala", "ID_Funcionario", "ID_LocalEvento", "Data", "Horario_Entrada", "Horario_Saida", "Status", "Funcao"]);
  }
  
  return _dbCache;
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

// Otimizado para usar a conexão em Cache
function getSheet(name) { 
  return getDB().getSheetByName(name); 
}

function sheetToJSON(data) {
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])))
    .filter(r => r[headers[0]] !== "");
}

function parseDateTime(dateStr, timeStr) {
  if (!dateStr || !timeStr) return null;
  const [y, m, d] = dateStr.split('-');
  const [hr, min] = timeStr.split(':');
  return new Date(y, m - 1, d, hr, min);
}

// ----------------------------------------------------------------------------
// 2. MOTORES DE NEGÓCIO E GET DATA
// ----------------------------------------------------------------------------

function getDashboardData() {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'dashboard_full_data';

    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    const locaisRaw = sheetToJSON(getSheet("LOCAIS_EVENTOS").getDataRange().getDisplayValues());
    const funcionarios = sheetToJSON(getSheet("FUNCIONARIOS").getDataRange().getDisplayValues());
    const escalasRaw = sheetToJSON(getSheet("ESCALAS").getDataRange().getDisplayValues());
    const funcoesRaw = sheetToJSON(getSheet("FUNCOES").getDataRange().getDisplayValues());
    
    // Funções padrão caso o BD seja novo
    let funcoes = funcoesRaw.length > 0 ? funcoesRaw : [
      {ID: '1', Nome: 'Vigilante'}, {ID: '2', Nome: 'Segurança Tático'}, {ID: '3', Nome: 'Chefe de Equipa'}
    ];

    const now = new Date();
    
    const locais = locaisRaw.map(local => {
      let expired = false;
      if (local.Status !== 'Cancelado' && local.Tipo === 'Evento' && local.Data_Fim && local.Hora_Fim) {
        const endDate = parseDateTime(local.Data_Fim, local.Hora_Fim);
        if (endDate && endDate < now) {
          expired = true;
          local.Status = 'Expirado';
        }
      }
      return { ...local, expired };
    });

    const escalas = escalasRaw.map(e => ({...e, Status: e.Status || 'Confirmado', Funcao: e.Funcao || 'Vigilante'}));

    const result = { success: true, data: { locais, funcionarios, escalas, funcoes } };
    // Cache por 5 minutos (300 segundos)
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(_) {}
    return result;
  } catch (e) { return { success: false, error: e.message }; }
}

// Validação de conflito baseada em Memória O(N) sem bater no App Script Quota
function validateConflictMemory(idFuncionario, dateStr, startStr, endStr, ignoreId, escalasCacheadas) {
  const newStart = parseDateTime(dateStr, startStr);
  let newEnd = parseDateTime(dateStr, endStr);
  if (!newStart || !newEnd) throw new Error("Data ou horário inválido.");
  if (newEnd <= newStart) newEnd.setDate(newEnd.getDate() + 1);

  for (let esc of escalasCacheadas) {
    if (esc.ID_Funcionario === idFuncionario && esc.Status !== 'Cancelado') {
      if (ignoreId && esc.ID_Escala === ignoreId) continue; 

      let escStart = parseDateTime(esc.Data, esc.Horario_Entrada);
      let escEnd = parseDateTime(esc.Data, esc.Horario_Saida);
      if (escEnd <= escStart) escEnd.setDate(escEnd.getDate() + 1);

      if (newStart < escEnd && newEnd > escStart) {
        throw new Error(`O agente selecionado já possui uma escala sobreposta no dia ${dateStr}.`);
      }
    }
  }
}

// Limpa o cache do dashboard após operações de escrita
function invalidateDashboardCache() {
  try {
    CacheService.getScriptCache().remove('dashboard_full_data');
  } catch(_) {}
}

// Cria índice O(1) de escalas por funcionário
function buildEscalasIndex(escalas) {
  const index = {};
  escalas.forEach(esc => {
    if (!index[esc.ID_Funcionario]) index[esc.ID_Funcionario] = [];
    index[esc.ID_Funcionario].push(esc);
  });
  return index;
}

// Validação de conflito otimizada usando índice O(1)
function validateConflictMemoryOptimized(idFuncionario, dateStr, startStr, endStr, ignoreId, escalasIndex) {
  const newStart = parseDateTime(dateStr, startStr);
  let newEnd = parseDateTime(dateStr, endStr);
  if (!newStart || !newEnd) throw new Error("Data ou horário inválido.");
  if (newEnd <= newStart) newEnd.setDate(newEnd.getDate() + 1);

  // O(1): acesso direto ao invés de O(N) loop sobre todas as escalas
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

// ----------------------------------------------------------------------------
// 3. API CRUD (CREATE, UPDATE, DELETE)
// ----------------------------------------------------------------------------

function updateOrAddRow(sheetName, idColIndex, id, newData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColIndex] === id) { rowIndex = i + 1; break; }
  }
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, newData.length).setValues([newData]); 
  } else {
    sheet.appendRow(newData); 
  }
}

function saveFuncao(nome) {
  try {
    const sheet = getSheet("FUNCOES");
    const id = Utilities.getUuid();
    sheet.appendRow([id, nome]);
    invalidateDashboardCache();
    return { success: true, data: { ID: id, Nome: nome } };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveFuncionario(payload) {
  try {
    const isNew = !payload.ID;
    const id = isNew ? Utilities.getUuid() : payload.ID;
    const row = [id, payload.Nome, payload.Telefone, payload.Status || 'Ativo'];
    updateOrAddRow("FUNCIONARIOS", 0, id, row);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveLocal(payload) {
  try {
    const isNew = !payload.ID;
    const id = isNew ? Utilities.getUuid() : payload.ID;
    const row = [id, payload.Nome, payload.Tipo, payload.Data_Inicio||'', payload.Hora_Inicio||'', payload.Data_Fim||'', payload.Hora_Fim||'', payload.Status || 'Ativo'];
    updateOrAddRow("LOCAIS_EVENTOS", 0, id, row);
    invalidateDashboardCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// Suporta Lote de Escalas (Lê apenas 1 vez, processa na velocidade da luz)
function saveEscalaBatch(payloads) {
  try {
    const sheet = getSheet("ESCALAS");
    
    // Ler todo o banco atual APENAS UMA VEZ
    const escalasAtuais = sheetToJSON(sheet.getDataRange().getDisplayValues());
    
    // OTIMIZAÇÃO: criar índice O(1) uma única vez antes de todas as validações
    const escalasIndex = buildEscalasIndex(escalasAtuais);

    // Validar conflitos na memória para TODOS os agentes antes de salvar qualquer um
    for (let payload of payloads) {
      validateConflictMemoryOptimized(payload.ID_Funcionario, payload.Data, payload.Horario_Entrada, payload.Horario_Saida, payload.ID_Escala, escalasIndex);
    }

    // Salvar todos (mantendo updateOrAddRow para compatibilidade)
    for (let payload of payloads) {
      const isNew = !payload.ID_Escala;
      const id = isNew ? Utilities.getUuid() : payload.ID_Escala;
      const row = [
        id, payload.ID_Funcionario, payload.ID_LocalEvento, 
        payload.Data, payload.Horario_Entrada, payload.Horario_Saida, 
        payload.Status || 'Confirmado', payload.Funcao || 'Vigilante'
      ];
      updateOrAddRow("ESCALAS", 0, id, row);
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
