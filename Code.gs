/*** ==============================================================================
* PROJET : COMPARATEUR D'INVENTAIRE - AVOSDIM
* VERSION : 6.9.10 (Live Stock Link for Validation)
* DATE : 26/12/2025
* ==============================================================================
*/

// --- CONFIGURATION ---
const DB_SPREADSHEET_ID = "1Gz-xX5YCaMQkaQBX83SJzKQ-l_D8_ARjEHu9ynPwcxo";
const LOG_SHEET_NAME = "LOG";
const MASTER_DB_SHEET_NAME = "BASE_ARTICLES";

// --- SERVLET & UTILS ---
function doGet() { return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Avosdim Stock v6.9.10').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function safeVal(v) { if (v === 0 || v === "0") return "0"; if (v === null || v === undefined || v === "") return ""; return String(v);}
function cleanSurstock(v) { if (v === null || v === undefined || v === "") return ""; let s = String(v).trim(); s = s.replace(/\//g, " | "); s = s.replace(/\s+/g, " | "); return s;}

// FORMATAGE DATE
function formatDateForDisplay(val) { 
  if (!val) return ""; 
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy"); 
  if (typeof val === 'string' && val.includes('-')) {
     const d = new Date(val);
     if(!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(val);
}
function parseDateLoose(value) { 
  if (!value) return null; 
  if (value instanceof Date) return value; 
  if (typeof value === 'string') { 
    if(value.includes('/')) {
        const parts = value.split('/');
        if (parts.length === 3) {
            const d = new Date(parts[2], parts[1]-1, parts[0]);
            if (!isNaN(d.getTime())) return d;
        }
    }
    const dIso = new Date(value); 
    if (!isNaN(dIso.getTime())) return dIso; 
  } 
  return null;
}
function safeDateString(v) { if (v instanceof Date && !isNaN(v)) return v.toISOString(); if (typeof v === 'string') return v; return "";}
function safeString(v) { return (v === null || v === undefined) ? "" : String(v); }
function safeNumber(v) { return (v === null || v === undefined || isNaN(v) || v === "") ? 0 : Number(v); }
function normalizeRef(v) { if (v === null || v === undefined) return ""; let s = String(v).trim(); if (/^\d+$/.test(s)) { return s.replace(/^0+/, '') || "0"; } return s;}

// Formatage quantité stricte (3 décimales)
function formatQty(val) {
    const n = Number(val);
    if (isNaN(n)) return 0;
    return Math.round(n * 1000) / 1000;
}

// --- MODULE 1 : BASE DE DONNEES ---
function getMasterDatabase() {
 try {
   const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
   const sheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
   if (!sheet) return { items: [], stats: { activeCount: 0, lastImport: "Aucun" } };
   
   let lastImportName = "Aucun";
   const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
   if (logSheet && logSheet.getLastRow() > 1) {
     const lastRowVal = logSheet.getRange(logSheet.getLastRow(), 2).getValue();
     lastImportName = (lastRowVal instanceof Date) ? Utilities.formatDate(lastRowVal, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(lastRowVal);
   }

   const dataRange = sheet.getDataRange();
   const rows = dataRange.getValues().slice(1);
   let activeCount = 0;
   const items = rows.map(r => {
     const isInactive = (r[25] !== "" && r[25] !== undefined); 
     if (!isInactive) activeCount++;
     const stockVal = (r.length > 28) ? r[28] : 0;
     return {
       ref: String(r[0]), name: String(r[1]), suppName: r[2], suppRef: r[3],
       rackF: r[4], rowF: r[5], caseF: r[6], surF: r[7],
       rackG: r[8], rowG: r[9], caseG: r[10], surG: r[11],
       rackA: r[12], rowA: r[13], caseA: r[14], surA: r[15],
       cote1: r[16], cote2: r[17], cote3: r[18], poids: r[19],
       dateAdd: formatDateForDisplay(r[20]), dateMod: formatDateForDisplay(r[21]), 
       lastSeen: formatDateForDisplay(parseDateLoose(r[22]) || r[22]), 
       lastCount: formatDateForDisplay(r[23]),
       dateDel: formatDateForDisplay(r[25]), 
       status: isInactive ? "Inactif" : "Actif",
       tags: r[26],
       stock: formatQty(stockVal)
     };
   });
   return { items: items, stats: { activeCount: activeCount, lastImport: lastImportName } };
 } catch (e) { throw new Error("Erreur BDD: " + e.message); }
}

function updateArticleTag(ref, newTagValue) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => String(r[0]) === String(ref));
  if (idx > -1) {
    sheet.getRange(idx + 1, 27).setValue(newTagValue);
    return { success: true };
  } else {
    return { success: false, error: "Référence introuvable" };
  }
}

function ensureSheetDimensions(sheet, requiredCols) {
  const currentCols = sheet.getMaxColumns();
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
}

function rebuildDatabase() {
 try {
   const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
   let backupCounts = new Map(); let backupTags = new Map();
   const existingSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
   if (existingSheet && existingSheet.getLastRow() > 1) {
       const exData = existingSheet.getDataRange().getValues();
       const headers = exData[0];
       const countIdx = headers.indexOf("Date Dernier Comptage");
       const tagsIdx = headers.indexOf("Tags");
       if(countIdx > -1) exData.slice(1).forEach(row => { if (row[countIdx]) backupCounts.set(normalizeRef(row[0]), row[countIdx]); });
       if(tagsIdx > -1) exData.slice(1).forEach(row => { if (row[tagsIdx]) backupTags.set(normalizeRef(row[0]), row[tagsIdx]); });
   }
   const history = getHistoryLog().reverse();
   const allSheets = ss.getSheets();
   const sheetMap = new Map();
   allSheets.forEach(s => sheetMap.set(s.getName(), s));
   let virtualDb = new Map();
   history.forEach(importLog => {
     try {
       const sheet = sheetMap.get(importLog.sheetName);
       if (sheet) {
         const values = sheet.getDataRange().getValues();
         const inventory = parseInventoryFromValues(values);
         let importDateObj = parseDateLoose(importLog.name) || new Date(importLog.date);
         if (isNaN(importDateObj.getTime())) importDateObj = new Date();
         processInventoryInMemory(virtualDb, inventory.data, importDateObj, importLog.name);
       }
     } catch (errFile) {}
   });
   virtualDb.forEach((item, refKey) => { 
       if (backupCounts.has(refKey)) item.lastCount = backupCounts.get(refKey); 
       if (backupTags.has(refKey)) item.tags = backupTags.get(refKey);
   });
   if (existingSheet) { ss.deleteSheet(existingSheet); }
   
   const dbSheet = ss.insertSheet(MASTER_DB_SHEET_NAME);
   ensureSheetDimensions(dbSheet, 29);
   dbSheet.getRange(1, 1, 5000, 29).setNumberFormat("@");
   
   const headers = ["Ref", "Nom", "Fournisseur", "Ref Fournisseur", "Rack F", "Row F", "Case F", "Surstock F", "Rack G", "Row G", "Case G", "Surstock G", "Rack A", "Row A", "Case A", "Surstock A", "Cote 1", "Cote 2", "Cote 3", "Poids", "Date Ajout", "Date Modif", "Dernière Vue (Tech)", "Date Dernier Comptage", "Temp", "Date Suppression", "Tags", "Date_Dernier_Inventaire", "Stock_Reel"];
   
   const outputData = [];
   virtualDb.forEach(item => {
     outputData.push([
       safeString(item.ref), safeString(item.name), safeString(item.suppName), safeString(item.suppRef), 
       safeVal(item.rackF), safeVal(item.rowF), safeVal(item.caseF), cleanSurstock(item.surF), 
       safeVal(item.rackG), safeVal(item.rowG), safeVal(item.caseG), cleanSurstock(item.surG), 
       safeVal(item.rackA), safeVal(item.rowA), safeVal(item.caseA), cleanSurstock(item.surA), 
       safeNumber(item.cote1), safeNumber(item.cote2), safeNumber(item.cote3), safeNumber(item.poids), 
       safeString(item.dateAdd), safeString(item.dateMod), 
       formatDateForDisplay(item.lastSeen), 
       safeString(item.lastCount), "", safeString(item.dateDel), safeString(item.tags), "",
       formatQty(item.qty)
     ]);
   });
   dbSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
   dbSheet.setFrozenRows(1);
   const CHUNK_SIZE = 1000;
   for (let i = 0; i < outputData.length; i += CHUNK_SIZE) {
       const chunk = outputData.slice(i, i + CHUNK_SIZE);
       if(chunk.length > 0) { dbSheet.getRange(2 + i, 1, chunk.length, headers.length).setValues(chunk); SpreadsheetApp.flush(); }
   }
   return `Base reconstruite.`;
 } catch (e) { throw new Error("Erreur: " + e.message); }
}

function processInventoryInMemory(dbMap, newInventory, importDateObj, dateStr) {
 const currentRefs = new Set();
 const safeImportDate = (importDateObj instanceof Date && !isNaN(importDateObj)) ? importDateObj : new Date();
 newInventory.forEach(item => {
   const ref = item.ref; currentRefs.add(ref);
   if (dbMap.has(ref)) {
     const existing = dbMap.get(ref);
     if (item.tags) existing.tags = item.tags;
     const lastSeenDate = parseDateLoose(existing.lastSeen) || new Date(0);
     if (safeImportDate >= lastSeenDate) {
       existing.lastSeen = safeImportDate;
       existing.qty = item.qty;
       existing.rackF = item.rackF; existing.rowF = item.rowF; existing.caseF = item.caseF; existing.surF = item.surF;
       existing.rackG = item.rackG; existing.rowG = item.rowG; existing.caseG = item.caseG; existing.surG = item.surG;
       existing.rackA = item.rackA; existing.rowA = item.rowA; existing.caseA = item.caseA; existing.surA = item.surA;
     }
   } else {
     dbMap.set(ref, { ref: ref, name: item.name, dateAdd: dateStr, lastSeen: safeImportDate, tags: item.tags || "", qty: item.qty });
   }
 });
}

function parseInventoryFromValues(values) {
 if (!values || values.length < 2) return { data: [], stockIdx: -1 };
 const normalizeHeader = (h) => String(h).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
 const headers = values[0].map(normalizeHeader);
 const data = values.slice(1);
 const parseNum = (val) => { 
    if (typeof val === 'string') {
        let clean = val.replace(/[\s\u00A0]/g, '').replace(',', '.');
        return parseFloat(clean) || 0; 
    } 
    return (typeof val === 'number') ? val : 0; 
 };
 let stockIdx = headers.findIndex(h => h === 'quantite en stock');
 if (stockIdx === -1) stockIdx = headers.findIndex(h => h === 'stock');
 if (stockIdx === -1) {
     stockIdx = headers.findIndex(h => h.includes('stock') && !h.includes('sur') && !h.includes('valeur') && !h.includes('alert'));
 }
 const col = (k) => headers.findIndex(h => h.includes(k.toLowerCase()));
 const rackF_idx = col('rack f') > -1 ? col('rack f') : col('rack');
 const colMap = { ref: col('ref'), stock: stockIdx, name: col('article/nom') > -1 ? col('article/nom') : col('nom'), suppName: col('fournisseurs/nom'), suppRef: col('fournisseurs/référence'), cote1: col('cote 1'), cote2: col('cote 2'), cote3: col('cote 3'), poids: col('poids'), rackF: rackF_idx, rowF: col('row f'), caseF: col('case f'), surF: col('surstock f'), rackG: col('rack g'), rowG: col('row g'), caseG: col('case g'), surG: col('surstock g'), rackA: col('rack a'), rowA: col('row a'), caseA: col('case a'), surA: col('surstock a'), tags: col('tags') > -1 ? col('tags') : col('etiquette') };
 const parsedData = data.map(row => {
   const getVal = (idx) => (idx > -1 && row[idx] !== undefined) ? row[idx] : "";
   const stockRaw = colMap.stock > -1 ? parseNum(row[colMap.stock]) : 0;
   return {
     ref: normalizeRef(String(row[colMap.ref])), name: colMap.name > -1 ? String(row[colMap.name]).trim() : "N/A", 
     qty: stockRaw,
     suppName: String(getVal(colMap.suppName)).trim(), suppRef: String(getVal(colMap.suppRef)).trim(),
     rackF: safeVal(getVal(colMap.rackF)), rowF: safeVal(getVal(colMap.rowF)), caseF: safeVal(getVal(colMap.caseF)), surF: safeVal(getVal(colMap.surF)),
     rackG: safeVal(getVal(colMap.rackG)), rowG: safeVal(getVal(colMap.rowG)), caseG: safeVal(getVal(colMap.caseG)), surG: safeVal(getVal(colMap.surG)),
     rackA: safeVal(getVal(colMap.rackA)), rowA: safeVal(getVal(colMap.rowA)), caseA: safeVal(getVal(colMap.caseA)), surA: safeVal(getVal(colMap.surA)),
     cote1: colMap.cote1 > -1 ? parseNum(row[colMap.cote1]) : 0, cote2: colMap.cote2 > -1 ? parseNum(row[colMap.cote2]) : 0, cote3: colMap.cote3 > -1 ? parseNum(row[colMap.cote3]) : 0, poids: colMap.poids > -1 ? parseNum(row[colMap.poids]) : 0,
     tags: safeString(getVal(colMap.tags))
   };
 }).filter(item => item.ref !== "");
 return { data: parsedData, stockIdx: stockIdx };
}

function updateMasterDbIncremental(ss, newInventory, importDateObj, importDateStr) {
 let dbSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
 const headers = ["Ref", "Nom", "Fournisseur", "Ref Fournisseur", "Rack F", "Row F", "Case F", "Surstock F", "Rack G", "Row G", "Case G", "Surstock G", "Rack A", "Row A", "Case A", "Surstock A", "Cote 1", "Cote 2", "Cote 3", "Poids", "Date Ajout", "Date Modif", "Dernière Vue (Tech)", "Date Dernier Comptage", "Temp", "Date Suppression", "Tags", "Date_Dernier_Inventaire", "Stock_Reel"];
 let dbMap = new Map();
 if (!dbSheet) {
   dbSheet = ss.insertSheet(MASTER_DB_SHEET_NAME);
   ensureSheetDimensions(dbSheet, headers.length);
   dbSheet.appendRow(headers);
   dbSheet.setFrozenRows(1);
   dbSheet.getRange("K:N").setNumberFormat("@");
 } else {
   ensureSheetDimensions(dbSheet, headers.length);
   dbSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
   const data = dbSheet.getDataRange().getValues();
   const dbData = data.slice(1);
   dbData.forEach((r) => {
     const val = (i) => (r.length > i && r[i] !== undefined) ? r[i] : "";
     const refKey = normalizeRef(val(0));
     dbMap.set(refKey, {
       ref: refKey, name: String(val(1)), suppName: val(2), suppRef: val(3),
       rackF: val(4), rowF: val(5), caseF: val(6), surF: val(7),
       rackG: val(8), rowG: val(9), caseG: val(10), surG: val(11),
       rackA: val(12), rowA: val(13), caseA: val(14), surA: val(15),
       cote1: val(16), cote2: val(17), cote3: val(18), poids: val(19),
       dateAdd: val(20), dateMod: val(21), lastSeen: parseDateLoose(val(22)) || new Date(0), lastCount: val(23), dateDel: val(25),
       tags: val(26), dateInv: val(27),
       qty: safeNumber(val(28))
     });
   });
 }
 processInventoryInMemory(dbMap, newInventory.data, importDateObj, importDateStr);
 const outputData = [];
 dbMap.forEach(item => {
   outputData.push([
       safeString(item.ref), safeString(item.name), safeString(item.suppName), safeString(item.suppRef),
       safeVal(item.rackF), safeVal(item.rowF), safeVal(item.caseF), cleanSurstock(item.surF),
       safeVal(item.rackG), safeVal(item.rowG), safeVal(item.caseG), cleanSurstock(item.surG),
       safeVal(item.rackA), safeVal(item.rowA), safeVal(item.caseA), cleanSurstock(item.surA),
       safeNumber(item.cote1), safeNumber(item.cote2), safeNumber(item.cote3), safeNumber(item.poids),
       safeString(item.dateAdd), safeString(item.dateMod), 
       formatDateForDisplay(item.lastSeen), 
       safeString(item.lastCount), "", safeString(item.dateDel), safeString(item.tags), safeString(item.dateInv),
       formatQty(item.qty)
   ]);
 });
 if (outputData.length > 0) {
   if(dbSheet.getLastRow() > 1) dbSheet.getRange(2, 1, dbSheet.getLastRow()-1, dbSheet.getLastColumn()).clearContent();
   const CHUNK_SIZE = 1000;
   for (let i = 0; i < outputData.length; i += CHUNK_SIZE) {
       const chunk = outputData.slice(i, i + CHUNK_SIZE);
       if(chunk.length > 0) { 
           dbSheet.getRange(2 + i, 1, chunk.length, headers.length).setValues(chunk); 
           SpreadsheetApp.flush(); 
       }
   }
 }
}

function saveImportToHistory(csvContent, userGivenName, fileName) {
 try {
   const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
   const separator = detectSeparator(csvContent);
   const csvData = Utilities.parseCsv(csvContent, separator);
   if (!csvData || csvData.length === 0) throw new Error("Fichier vide");
   let sheetName = userGivenName.replace(/\//g, "-");
   let sheet = ss.getSheetByName(sheetName);
   if (sheet) sheetName += "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmm");
   sheet = ss.insertSheet(sheetName);
   if (csvData.length > 0) {
       const rows = csvData.length; const cols = csvData[0].length;
       sheet.getRange(1, 1, rows, cols).setNumberFormat("@");
       sheet.getRange(1, 1, rows, cols).setValues(csvData);
   }
   let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
   if (!logSheet) { logSheet = ss.insertSheet(LOG_SHEET_NAME); logSheet.appendRow(["ID Session", "Nom Session", "Date Import", "Nom Onglet GSheet"]); }
   const sessionId = Date.now().toString();
   const importDateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
   const parseResult = parseInventoryFromValues(csvData);
   const debugInfo = (parseResult.stockIdx > -1) ? `Stock Col: ${parseResult.stockIdx}` : "STOCK COL NOT FOUND";
   logSheet.appendRow([sessionId, userGivenName, importDateStr, sheetName + " | " + debugInfo]);
   let currentImportDate = parseDateLoose(userGivenName) || new Date();
   updateMasterDbIncremental(ss, parseResult, currentImportDate, userGivenName);
   SpreadsheetApp.flush();
   return { success: true, message: "Sauvegardé (" + debugInfo + ")", id: sessionId };
 } catch (e) { throw new Error(e.message); }
}
function getHistoryLog() { try { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const logSheet = ss.getSheetByName(LOG_SHEET_NAME); if (!logSheet) return []; const data = logSheet.getDataRange().getValues(); if (data.length <= 1) return []; const history = []; for (let i = 1; i < data.length; i++) { const row = data[i]; if (row[0] || row[1]) { let name = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(row[1]); let dateStr = row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(row[2]); let sheetNameRaw = row[3]; let sheetName; if (sheetNameRaw instanceof Date) { sheetName = Utilities.formatDate(sheetNameRaw, Session.getScriptTimeZone(), "dd-MM-yyyy"); } else if (sheetNameRaw) { sheetName = String(sheetNameRaw); } else { sheetName = name.replace(/\//g, "-"); } history.push({ id: String(row[0]), name: name, date: dateStr, sheetName: sheetName.trim() }); } } return history.reverse(); } catch (e) { throw new Error("Erreur lecture LOG : " + e.message); } }
function detectSeparator(csvString) { const firstLine = csvString.split('\n')[0]; return (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ','; }
function checkDbConnection() { try { SpreadsheetApp.openById(DB_SPREADSHEET_ID); return true; } catch (e) { return false; } }

// --- MODULE 2 : AUTH ---
function getPublicUserList() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config_Users');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues().slice(1);
  return data
    .filter(r => String(r[6]).toLowerCase() === 'true' || r[6] === true)
    .map(r => ({ id: r[0], displayName: `${r[2]} ${r[1]}`, role: r[4] }));
}
function loginUser(userId, pinInput) {
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

// --- MODULE 3 : TACHES & GENERATION ---

// TRIGGER MENSUEL
function installMonthlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if(t.getHandlerFunction() === 'autoGenerateTasks') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoGenerateTasks').timeBased().onMonthDay(1).atHour(6).create();
  return { success: true, message: "Automatisation activée : 1er du mois à 06h00." };
}

function autoGenerateTasks() {
  const settings = getSettingsData();
  const batchSize = settings.params.find(p => p.key === 'AUTO_BATCH_SIZE');
  const size = batchSize ? Number(batchSize.value) : 50;
  generateTasksBatch(size, 'all');
}

function resetAndGenerateTasks(maxItems, zoneFilter) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if (taskSheet.getLastRow() > 1) taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).clearContent();
  return generateTasksBatch(maxItems, zoneFilter);
}

function generateTasksBatch(maxItems, zoneFilter) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const articlesSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const articles = articlesSheet.getDataRange().getValues().slice(1);
  const rules = getTagsRules();
  const today = new Date();
  const candidates = [];
  
  for(let row of articles) {
      if(row[25] !== "") continue; 
      let hasLoc = false;
      if (zoneFilter === 'all' || !zoneFilter) hasLoc = true;
      else if (zoneFilter === 'F' && (row[4] || row[7])) hasLoc = true;
      else if (zoneFilter === 'G' && (row[8] || row[11])) hasLoc = true;
      else if (zoneFilter === 'A' && (row[12] || row[15])) hasLoc = true;
      if(!hasLoc) continue;

      const tagsStr = String(row[26] || "").toUpperCase();
      const lastInvRaw = row[27];
      let freq = rules.get('DEFAULT');
      if (tagsStr) {
          const tags = tagsStr.replace(/[\[\]"]/g, '').split(/[ ,]+/);
          let minFreq = 9999; let foundTag = false;
          tags.forEach(t => { if(rules.has(t)) { const f = rules.get(t); if(f < minFreq) minFreq = f; foundTag = true; } });
          if(foundTag) freq = minFreq;
      }
      let lastDate = parseDateLoose(lastInvRaw);
      if(!lastDate) lastDate = new Date(2000, 0, 1);
      const nextDueDate = new Date(lastDate); nextDueDate.setDate(lastDate.getDate() + freq);
      if (nextDueDate <= today) {
          const diffTime = Math.abs(today - nextDueDate);
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
          candidates.push({ row: row, score: diffDays });
      }
  }
  
  candidates.sort((a,b) => b.score - a.score);
  const selected = candidates.slice(0, maxItems);
  
  const newTasks = selected.map(item => {
      const row = item.row;
      const formatLoc = (label, r, ro, c, s) => {
          let p = [];
          const rv = safeVal(r); if(rv !== "") p.push(rv);
          const rov = safeVal(ro); if(rov !== "") p.push(rov);
          const cv = safeVal(c); if(cv !== "") p.push(cv);
          if (p.length === 0 && !s) return null;
          let str = p.join('-');
          const sv = cleanSurstock(s);
          if(sv !== "") str += ` (+${sv})`;
          return str ? `${label}: ${str}` : null;
      };
      let locs = [];
      const lF = formatLoc('F', row[4], row[5], row[6], row[7]);
      const lG = formatLoc('G', row[8], row[9], row[10], row[11]);
      const lA = formatLoc('A', row[12], row[13], row[14], row[15]);
      if(lF) locs.push(lF); if(lG) locs.push(lG); if(lA) locs.push(lA);
      const locStr = locs.join(' || ');
      
      // RECUPERATION STOCK (Index 28 = Column 29)
      const stockTheorique = (row.length > 28) ? safeNumber(row[28]) : 0;
      
      return [Utilities.getUuid(), String(row[0]), locStr, stockTheorique, new Date(), "", "A_FAIRE", "", "", ""];
  });
  
  if(newTasks.length > 0) taskSheet.appendRow(newTasks[0]); 
  if(newTasks.length > 0) {
      newTasks.forEach(r => taskSheet.appendRow(r)); 
  }
  return { count: newTasks.length };
}

function getTagsRules() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config_Settings');
  if(!sheet) return new Map();
  const data = sheet.getDataRange().getValues().slice(1);
  const rules = new Map();
  rules.set('DEFAULT', 180); 
  data.forEach(r => { if(r[0] === 'TAG' && r[2]) rules.set(String(r[1]).toUpperCase(), Number(r[2])); });
  return rules;
}

// --- GESTION LISTE (ADD/REMOVE) ---

function addTaskManually(ref) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const data = articleSheet.getDataRange().getValues();
  const row = data.find(r => String(r[0]) === String(ref));
  if (!row) return { success: false, error: "Référence inconnue" };
  
  const formatLoc = (label, r, ro, c, s) => {
      let p = [];
      const rv = safeVal(r); if(rv !== "") p.push(rv);
      const rov = safeVal(ro); if(rov !== "") p.push(rov);
      const cv = safeVal(c); if(cv !== "") p.push(cv);
      if (p.length === 0 && !s) return null;
      let str = p.join('-');
      const sv = cleanSurstock(s);
      if(sv !== "") str += ` (+${sv})`;
      return str ? `${label}: ${str}` : null;
  };
  let locs = [];
  const lF = formatLoc('F', row[4], row[5], row[6], row[7]);
  const lG = formatLoc('G', row[8], row[9], row[10], row[11]);
  const lA = formatLoc('A', row[12], row[13], row[14], row[15]);
  if(lF) locs.push(lF); if(lG) locs.push(lG); if(lA) locs.push(lA);
  const locStr = locs.join(' || '); 
  
  const stockTheorique = (row.length > 28) ? safeNumber(row[28]) : 0;
  
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  taskSheet.appendRow([Utilities.getUuid(), String(row[0]), locStr, stockTheorique, new Date(), "MANUEL", "A_FAIRE", "", "", ""]);
  return { success: true };
}

function deleteTask(taskId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Data_CycleTasks');
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => String(r[0]) === String(taskId));
  if (idx > 0) { 
    sheet.deleteRow(idx + 1);
    return { success: true };
  }
  return { success: false, error: "Tâche introuvable" };
}

function getInProgressTasks() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  if (!taskSheet) return [];
  const data = taskSheet.getDataRange().getValues().slice(1);
  return data
    .filter(r => r[6] === 'A_FAIRE' || r[6] === 'EN_COURS')
    .map(r => ({ id: r[0], ref: r[1], loc: r[2], status: r[6] }));
}

function getOperatorTodoList(userId, zoneFilter) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const articleSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const tasks = taskSheet.getDataRange().getValues().slice(1).filter(r => r[6] === 'A_FAIRE' || (r[6] === 'EN_COURS' && String(r[5]) === String(userId)));
  if(tasks.length === 0) return [];
  const articleData = articleSheet.getDataRange().getValues();
  const articleMap = new Map();
  articleData.slice(1).forEach(r => articleMap.set(String(r[0]), { name: r[1] }));
  const todoList = tasks.map(t => {
      const art = articleMap.get(String(t[1])) || { name: "Inconnu" };
      let locDisplay = t[2];
      if (zoneFilter && zoneFilter !== 'all') {
          const parts = t[2].split(' || ');
          const match = parts.find(p => p.trim().startsWith(zoneFilter + ':'));
          if (match) locDisplay = match.substring(2).trim(); 
      }
      return { taskId: t[0], ref: t[1], locSnap: locDisplay, name: art.name, status: t[6] };
  });
  todoList.sort((a,b) => a.locSnap.localeCompare(b.locSnap));
  return todoList;
}

function submitTask(taskId, qty, userId) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const data = taskSheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => String(r[0]) === String(taskId));
  if(rowIndex === -1) return { success: false, error: "Tâche introuvable" };
  const rowNum = rowIndex + 1;
  taskSheet.getRange(rowNum + 1, 6).setValue(userId);
  taskSheet.getRange(rowNum + 1, 7).setValue('COMPTE');
  taskSheet.getRange(rowNum + 1, 8).setValue(qty);
  taskSheet.getRange(rowNum + 1, 9).setValue(new Date());
  return { success: true };
}

// --- MODULE 4 : VALIDATION ---
function getPendingValidations() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  
  // 1. Charger le stock réel actuel depuis la BDD (Fix v6.9.10)
  const artSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const artData = artSheet.getDataRange().getValues();
  const stockMap = new Map();
  // Skip headers
  for(let i=1; i<artData.length; i++) {
      // Ref is index 0, Stock_Reel is index 28
      const ref = String(artData[i][0]);
      // Protection lecture si colonne 28 n'existe pas encore pour cette ligne
      const stock = (artData[i].length > 28) ? safeNumber(artData[i][28]) : 0;
      stockMap.set(ref, stock);
  }

  // 2. Charger les tâches
  const taskSheet = ss.getSheetByName('Data_CycleTasks');
  const tasks = taskSheet.getDataRange().getValues().slice(1);
  const pending = tasks.filter(r => r[6] === 'COMPTE');
  
  return pending.map(r => {
      const ref = String(r[1]);
      // On privilégie le stock en direct de la BDD, sinon on prend le snapshot (col 3)
      const currentStock = stockMap.has(ref) ? stockMap.get(ref) : safeNumber(r[3]);
      const counted = safeNumber(r[7]);
      
      // On force le format 3 décimales pour l'affichage
      const theoFormatted = formatQty(currentStock);
      
      return { 
          taskId: r[0], 
          ref: ref, 
          loc: r[2], 
          theo: theoFormatted, 
          counted: counted, 
          delta: formatQty(counted - theoFormatted), // Recalcul du delta
          user: r[5], 
          date: formatDateForDisplay(r[8]) 
      };
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

// --- MODULE 5 : SETTINGS ---
function getSettingsData() {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const userSheet = ss.getSheetByName('Config_Users');
  const users = userSheet.getDataRange().getValues().slice(1).map(r => ({ id: r[0], nom: r[1], prenom: r[2], pin: r[3], role: r[4], zone: r[5], actif: r[6] }));
  const settingSheet = ss.getSheetByName('Config_Settings');
  const settingsData = settingSheet.getDataRange().getValues().slice(1);
  const tags = settingsData.filter(r => r[0] === 'TAG').map(r => ({ name: r[1], freq: r[2], color: r[3], desc: r[4] }));
  const params = settingsData.filter(r => r[0] === 'PARAM').map(r => ({ key: r[1], value: r[2], desc: r[4] }));
  return { users, tags, params };
}
function saveUser(user) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config_Users');
  const data = sheet.getDataRange().getValues();
  if (user.id) { const idx = data.findIndex(r => String(r[0]) === String(user.id)); if (idx > 0) { const r = idx + 1; sheet.getRange(r, 2).setValue(user.nom); sheet.getRange(r, 3).setValue(user.prenom); sheet.getRange(r, 4).setValue(user.pin); sheet.getRange(r, 5).setValue(user.role); sheet.getRange(r, 6).setValue(user.zone); } } else { const newId = 'U' + Date.now(); sheet.appendRow([newId, user.nom, user.prenom, user.pin, user.role, user.zone, true, '']); }
  return getSettingsData();
}
function deleteUser(id) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Users'); const data = sheet.getDataRange().getValues(); const idx = data.findIndex(r => String(r[0]) === String(id)); if (idx > 0) sheet.deleteRow(idx + 1); return getSettingsData(); }
function saveTag(tag) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sheet = ss.getSheetByName('Config_Settings'); const data = sheet.getDataRange().getValues(); const idx = data.findIndex(r => r[0] === 'TAG' && r[1] === tag.name); if (idx > 0) { sheet.getRange(idx + 1, 3).setValue(tag.freq); sheet.getRange(idx + 1, 4).setValue(tag.color); } return getSettingsData(); }
function saveParam(key, val) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Config_Settings');
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => r[0] === 'PARAM' && r[1] === key);
  if (idx > -1) { sheet.getRange(idx + 1, 3).setValue(val); return { success: true }; }
  // Create if not exist
  sheet.appendRow(['PARAM', key, val, '', 'Paramètre généré']);
  return { success: true };
}

// --- MODULE 6 : HISTORY ---
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
