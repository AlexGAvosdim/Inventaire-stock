// --- MODULE APP & USERS ---

function getPublicUserList() {
  try {
    const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Config_Users');
    if (!sheet) return [{ id: 'rescue_admin', displayName: '⚠️ Admin Secours (0000)', role: 'Admin' }];
    if (sheet.getLastRow() <= 1) return [{ id: 'rescue_admin', displayName: '⚠️ Admin Secours (0000)', role: 'Admin' }];
    const data = sheet.getDataRange().getValues().slice(1);
    const users = data.filter(r => r.length > 6 && (String(r[6]).toLowerCase() === 'true' || r[6] === true)).map(r => ({ id: r[0], displayName: `${r[2]} ${r[1]}`, role: r[4] }));
    if (users.length === 0) return [{ id: 'rescue_admin', displayName: '⚠️ Admin Secours (0000)', role: 'Admin' }];
    return users;
  } catch (e) { return [{ id: 'rescue_admin', displayName: '⚠️ Admin Secours (0000)', role: 'Admin' }]; }
}

function loginUser(userId, pinInput) {
  if (userId === 'rescue_admin') {
      if (pinInput === '0000') return { success: true, user: { id: 'rescue_admin', nom: 'Admin', prenom: 'Secours', role: 'Admin', zone: 'All' } };
      else return { success: false, error: "Code PIN incorrect" };
  }
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config_Users');
  const data = sheet.getDataRange().getValues();
  const userRowIndex = data.findIndex(r => String(r[0]) === String(userId));
  if (userRowIndex === -1) return { success: false, error: "Utilisateur introuvable" };
  const row = data[userRowIndex];
  if (String(row[3]) === String(pinInput)) {
    const user = { id: row[0], nom: row[1], prenom: row[2], role: row[4], zone: row[5] };
    sheet.getRange(userRowIndex + 1, 8).setValue(new Date());
    return { success: true, user: user };
  } else { return { success: false, error: "Code PIN incorrect" }; }
}

function getSettingsData() {
  try {
    const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
    let users = [], tags = [], params = [];
    try {
        const userSheet = ss.getSheetByName('Config_Users');
        if (userSheet && userSheet.getLastRow() > 1) { const uData = userSheet.getDataRange().getValues(); users = uData.slice(1).map(r => ({ id: r[0], nom: r[1], prenom: r[2], pin: r[3], role: r[4], zone: r[5], actif: r[6] })); }
    } catch(e) {}
    try {
        const settingSheet = ss.getSheetByName('Config_Settings');
        if (settingSheet && settingSheet.getLastRow() > 1) { const sData = settingSheet.getDataRange().getValues().slice(1); tags = sData.filter(r => r[0] === 'TAG').map(r => ({ name: r[1], freq: r[2], color: r[3], desc: r[4] })); params = sData.filter(r => r[0] === 'PARAM').map(r => ({ key: r[1], value: r[2], desc: r[4] })); }
    } catch(e) {}
    let triggerActive = false;
    try { triggerActive = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoGenerateTasks'); } catch(e) {}
    return { users, tags, params, triggerActive };
  } catch (e) { throw new Error("Erreur getSettingsData: " + e.message); }
}

function toggleAutoTrigger(enable) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => { if(t.getHandlerFunction() === 'autoGenerateTasks') ScriptApp.deleteTrigger(t); });
    if (enable) { ScriptApp.newTrigger('autoGenerateTasks').timeBased().onMonthDay(1).atHour(6).create(); return { success: true, message: "✅ Activé" }; } else { return { success: true, message: "❌ Désactivé" }; }
  } catch(e) { throw e; }
}

function autoGenerateTasks() { const settings = getSettingsData(); const batchSize = settings.params.find(p => p.key === 'AUTO_BATCH_SIZE'); const size = batchSize ? Number(batchSize.value) : 50; generateTasksBatch(size, 'all'); }

function resetAndGenerateTasks(maxItems, zoneFilter, keepExisting) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if (!keepExisting && taskSheet.getLastRow() > 1) {
       if (zoneFilter === 'all') { taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).clearContent(); } 
       else {
           const data = taskSheet.getDataRange().getValues();
           const rowsKeep = [];
           for(let i=1; i<data.length; i++) { if (!String(data[i][2]).startsWith(zoneFilter + ':')) rowsKeep.push(data[i]); }
           taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).clearContent();
           if (rowsKeep.length > 0) taskSheet.getRange(2, 1, rowsKeep.length, rowsKeep[0].length).setValues(rowsKeep);
       }
  }
  return generateTasksBatch(maxItems, zoneFilter);
}

function saveUser(user) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Users'); const data = sheet.getDataRange().getValues(); if (user.id) { const idx = data.findIndex(r => String(r[0]) === String(user.id)); if (idx > 0) { const r = idx + 1; sheet.getRange(r, 2).setValue(user.nom); sheet.getRange(r, 3).setValue(user.prenom); sheet.getRange(r, 4).setValue(user.pin); sheet.getRange(r, 5).setValue(user.role); sheet.getRange(r, 6).setValue(user.zone); } } else { const newId = 'U' + Date.now(); sheet.appendRow([newId, user.nom, user.prenom, user.pin, user.role, user.zone, true, '']); } return getSettingsData(); }
function deleteUser(id) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Users'); const data = sheet.getDataRange().getValues(); const idx = data.findIndex(r => String(r[0]) === String(id)); if (idx > 0) sheet.deleteRow(idx + 1); return getSettingsData(); }
function saveTag(tag) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Settings'); const data = sheet.getDataRange().getValues(); const idx = data.findIndex(r => r[0] === 'TAG' && r[1] === tag.name); if (idx > 0) { sheet.getRange(idx + 1, 3).setValue(tag.freq); sheet.getRange(idx + 1, 4).setValue(tag.color); } return getSettingsData(); }
function saveParam(key, val) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Settings'); const data = sheet.getDataRange().getValues(); const idx = data.findIndex(r => r[0] === 'PARAM' && r[1] === key); if (idx > -1) { sheet.getRange(idx + 1, 3).setValue(val); return { success: true }; } sheet.appendRow(['PARAM', key, val, '', 'Paramètre généré']); return { success: true }; }

function getTaskHistory() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if(!taskSheet) return [];
  const tasks = taskSheet.getDataRange().getValues().slice(1);
  const history = tasks.filter(r => r[6] === 'VALIDE');
  const userSheet = ss.getSheetByName('Config_Users');
  const userMap = new Map();
  if(userSheet) userSheet.getDataRange().getValues().slice(1).forEach(r => userMap.set(String(r[0]), `${r[2]} ${r[1]}`));
  return history.slice(-100).reverse().map(r => {
      const theo = safeNumber(r[3]); const counted = safeNumber(r[7]);
      return { date: formatDateForDisplay(r[8]), ref: r[1], loc: r[2], user: userMap.get(String(r[5])) || r[5], theo: theo, counted: counted, delta: counted - theo };
  });
}

function getOperatorTodoList(userId, zoneFilter) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if(!taskSheet) return [];
  const tasks = taskSheet.getDataRange().getValues().slice(1);
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const articleData = articleSheet.getDataRange().getValues();
  const articleMap = new Map();
  articleData.slice(1).forEach(r => articleMap.set(String(r[0]), { name: r[1] }));
  const filtered = tasks.filter(r => {
      const statusOk = r[6] === 'A_FAIRE' || (r[6] === 'EN_COURS' && String(r[5]) === String(userId));
      if (!statusOk) return false;
      const loc = String(r[2]);
      if (zoneFilter && zoneFilter !== 'all' && zoneFilter !== 'All') { const prefix = zoneFilter.toUpperCase() + ':'; if (!loc.toUpperCase().startsWith(prefix)) return false; }
      return true;
  });
  const todoList = filtered.map(t => {
      const art = articleMap.get(String(t[1])) || { name: "Inconnu" };
      const priority = (t.length > 10) ? Number(t[10]) : 0;
      const type = String(t[5]); 
      // V7.2.0: ADD DATE FROM COL 4
      const dateGen = formatDateForDisplay(t[4]);
      return { taskId: t[0], ref: t[1], locSnap: t[2], name: art.name, status: t[6], prio: priority, type: type, date: dateGen };
  });
  todoList.sort((a,b) => { if (a.prio !== b.prio) return b.prio - a.prio; return a.locSnap.localeCompare(b.locSnap); });
  return todoList;
}

function getInProgressTasks() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if (!taskSheet) return [];
  const data = taskSheet.getDataRange().getValues().slice(1);
  const pending = data.filter(r => r[6] === 'A_FAIRE' || r[6] === 'EN_COURS');
  if (pending.length === 0) return [];
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const articleData = articleSheet.getDataRange().getValues();
  const articleMap = new Map();
  articleData.slice(1).forEach(r => articleMap.set(String(r[0]), { name: r[1], stock: (r.length > 28) ? r[28] : 0 }));
  return pending.map(r => {
      const ref = String(r[1]);
      const art = articleMap.get(ref) || { name: "Inconnu", stock: 0 };
      const locStr = String(r[2]);
      let zone = "?"; if(locStr.startsWith("F:")) zone = "F"; else if(locStr.startsWith("G:")) zone = "G"; else if(locStr.startsWith("A:")) zone = "A";
      const type = String(r[5]); const prio = (r.length > 10) ? Number(r[10]) : 0;
      // V7.2.0: ADD DATE
      const dateGen = formatDateForDisplay(r[4]);
      return { taskId: r[0], ref: ref, name: art.name, loc: locStr, zone: zone, status: r[6], user: (r[6] === 'EN_COURS') ? r[5] : "", type: type, prio: prio, stockTheo: formatQty(art.stock), date: dateGen };
  });
}

function submitTask(taskId, qty, userId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const data = taskSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => String(r[0]) === String(taskId));
  if(rowIndex === -1) return { success: false, error: "Tâche introuvable" };
  const rowNum = rowIndex + 1;
  taskSheet.getRange(rowNum, 6).setValue(userId);
  taskSheet.getRange(rowNum, 7).setValue('COMPTE');
  taskSheet.getRange(rowNum, 8).setValue(qty);
  taskSheet.getRange(rowNum, 9).setValue(new Date());
  return { success: true };
}

function getPendingValidations() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const artSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const artData = artSheet.getDataRange().getValues();
  const stockMap = new Map();
  const nameMap = new Map();
  for(let i=1; i<artData.length; i++) { const ref = String(artData[i][0]); stockMap.set(ref, (artData[i].length > 28) ? safeNumber(artData[i][28]) : 0); nameMap.set(ref, String(artData[i][1])); }
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const tasks = taskSheet.getDataRange().getValues().slice(1);
  const pendingSiblings = {};
  tasks.forEach(r => {
      if (r[6] === 'A_FAIRE' || r[6] === 'EN_COURS') {
          const ref = String(r[1]);
          if(!pendingSiblings[ref]) pendingSiblings[ref] = [];
          const loc = String(r[2]);
          const zone = loc.split(':')[0];
          pendingSiblings[ref].push({ id: r[0], zone: zone, prio: (r.length > 10 ? r[10] : 0), type: String(r[5]) });
      }
  });
  const toValidate = tasks.filter(r => r[6] === 'COMPTE');
  return toValidate.map(r => {
      const ref = String(r[1]);
      const currentStock = stockMap.has(ref) ? stockMap.get(ref) : safeNumber(r[3]);
      const counted = safeNumber(r[7]);
      const siblings = pendingSiblings[ref] || [];
      const locStr = String(r[2]);
      let zone = "?"; if(locStr.startsWith("F:")) zone = "F"; else if(locStr.startsWith("G:")) zone = "G"; else if(locStr.startsWith("A:")) zone = "A";
      return { taskId: r[0], ref: ref, name: nameMap.get(ref) || "Inconnu", loc: locStr, zone: zone, theo: formatQty(currentStock), counted: counted, user: r[5], date: formatDateForDisplay(r[8]), others: siblings, delta: counted - currentStock };
  });
}

function validateTaskBatch(taskIds) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const tasksData = taskSheet.getDataRange().getValues();
  const articleData = articleSheet.getDataRange().getValues();
  const articleMap = new Map();
  for(let i=1; i<articleData.length; i++) articleMap.set(String(articleData[i][0]), i+1);
  let count = 0; const today = new Date();
  taskIds.forEach(id => {
      const idx = tasksData.findIndex(r => String(r[0]) === String(id));
      if (idx > 0) {
          const rowNum = idx + 1; const ref = String(tasksData[idx][1]);
          taskSheet.getRange(rowNum, 7).setValue('VALIDE');
          taskSheet.getRange(rowNum, 10).setValue(today);
          if (articleMap.has(ref)) { articleSheet.getRange(articleMap.get(ref), 28).setValue(today); }
          count++;
      }
  });
  return { success: true, count: count };
}

function deleteTask(taskId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_CycleTasks');
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => String(r[0]) === String(taskId));
  if (idx > 0) { sheet.deleteRow(idx + 1); return { success: true }; }
  return { success: false, error: "Tâche introuvable" };
}

function rejectTask(taskId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_CycleTasks');
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(r => String(r[0]) === String(taskId));
  if (rowIdx > -1) {
    const r = rowIdx + 1;
    sheet.getRange(r, 6).setValue("RECOMPTAGE"); 
    sheet.getRange(r, 7).setValue("A_FAIRE");    
    sheet.getRange(r, 8).clearContent();         
    sheet.getRange(r, 9).clearContent();         
    sheet.getRange(r, 11).setValue(1);           
    return { success: true };
  }
  return { success: false, error: "Tâche introuvable" };
}

function toggleTaskPriority(taskId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_CycleTasks');
  ensureSheetDimensions(sheet, 11);
  const data = sheet.getDataRange().getValues();
  const rowIdx = data.findIndex(r => String(r[0]) === String(taskId));
  if (rowIdx > -1) {
    const current = Number(data[rowIdx][10]) || 0;
    const newVal = current === 1 ? 0 : 1;
    sheet.getRange(rowIdx + 1, 11).setValue(newVal);
    return { success: true, newVal: newVal };
  }
  return { success: false, error: "Tâche introuvable" };
}

function addTaskManually(ref, isPriority) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const data = articleSheet.getDataRange().getValues();
  const row = data.find(r => String(r[0]) === String(ref));
  if (!row) return { success: false, error: "Référence inconnue" };
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  ensureSheetDimensions(taskSheet, 11);
  const formatLoc = (label, r, ro, c, s) => { let p = []; const rv = safeVal(r); if(rv !== "") p.push(rv); const rov = safeVal(ro); if(rov !== "") p.push(rov); const cv = safeVal(c); if(cv !== "") p.push(cv); if (p.length === 0 && !s) return null; let str = p.join('-'); const sv = cleanSurstock(s); if(sv !== "") str += ` (+${sv})`; return `${label}: ${str}`; };
  const lF = formatLoc('F', row[4], row[5], row[6], row[7]);
  const lG = formatLoc('G', row[8], row[9], row[10], row[11]);
  const lA = formatLoc('A', row[12], row[13], row[14], row[15]);
  const stock = (row.length > 28) ? safeNumber(row[28]) : 0;
  const prioVal = isPriority ? 1 : 0;
  const newTasks = [];
  if (lF) newTasks.push([Utilities.getUuid(), String(row[0]), lF, stock, new Date(), "MANUEL", "A_FAIRE", "", "", "", prioVal]);
  if (lG) newTasks.push([Utilities.getUuid(), String(row[0]), lG, stock, new Date(), "MANUEL", "A_FAIRE", "", "", "", prioVal]);
  if (lA) newTasks.push([Utilities.getUuid(), String(row[0]), lA, stock, new Date(), "MANUEL", "A_FAIRE", "", "", "", prioVal]);
  if (newTasks.length === 0) newTasks.push([Utilities.getUuid(), String(row[0]), "F: (Manuel)", stock, new Date(), "MANUEL", "A_FAIRE", "", "", "", prioVal]);
  taskSheet.getRange(taskSheet.getLastRow() + 1, 1, newTasks.length, newTasks[0].length).setValues(newTasks);
  return { success: true };
}
