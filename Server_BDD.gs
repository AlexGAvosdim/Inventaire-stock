// --- MODULE BDD ---

function getMasterDatabase() {
 try {
   const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
   const sheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
   if (!sheet) return { items: [], stats: { activeCount: 0, lastImport: "Aucun" } };
   
   let lastImportName = "Aucun";
   try {
     const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
     if (logSheet && logSheet.getLastRow() > 1) {
       lastImportName = formatDateForDisplay(logSheet.getRange(logSheet.getLastRow(), 2).getValue());
     }
   } catch(e) {}

   const dataRange = sheet.getDataRange();
   const rows = dataRange.getValues().slice(1);
   let activeCount = 0;
   
   const items = rows.map(r => {
     const isInactive = (r[25] !== ""); 
     if (!isInactive) activeCount++;
     return {
       ref: safeString(r[0]), name: safeString(r[1]), suppName: safeString(r[2]), suppRef: safeString(r[3]),
       rackF: r[4], rowF: r[5], caseF: r[6], surF: r[7],
       rackG: r[8], rowG: r[9], caseG: r[10], surG: r[11],
       rackA: r[12], rowA: r[13], caseA: r[14], surA: r[15],
       cote1: r[16], cote2: r[17], cote3: r[18], poids: r[19],
       dateAdd: formatDateForDisplay(r[20]), dateMod: formatDateForDisplay(r[21]), 
       lastSeen: formatDateForDisplay(parseDateLoose(r[22]) || r[22]), 
       lastCount: formatDateForDisplay(r[23]),
       dateDel: formatDateForDisplay(r[25]), 
       status: isInactive ? "Inactif" : "Actif",
       tags: safeString(r[26]),
       stock: formatQty(r[28])
     };
   });
   return { items: items, stats: { activeCount: activeCount, lastImport: lastImportName } };
 } catch (e) { throw new Error("Erreur BDD: " + e.message); }
}

function updateArticleTag(ref, newTag) {
  const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_DB_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => String(r[0]) === String(ref));
  if (idx > -1) { sheet.getRange(idx + 1, 27).setValue(newTag); return { success: true }; }
  return { success: false, error: "Référence introuvable" };
}

function updateMasterDbIncremental(ss, newInventory, importDateObj, importDateStr) { 
  let dbSheet = ss.getSheetByName(MASTER_DB_SHEET_NAME); 
  const headers = ["Ref", "Nom", "Fournisseur", "Ref Fournisseur", "Rack F", "Row F", "Case F", "Surstock F", "Rack G", "Row G", "Case G", "Surstock G", "Rack A", "Row A", "Case A", "Surstock A", "Cote 1", "Cote 2", "Cote 3", "Poids", "Date Ajout", "Date Modif", "Dernière Vue (Tech)", "Date Dernier Comptage", "Temp", "Date Suppression", "Tags", "Date_Dernier_Inventaire", "Stock_Reel"]; 
  let dbMap = new Map(); 
  if (!dbSheet) { 
    dbSheet = ss.insertSheet(MASTER_DB_SHEET_NAME); dbSheet.appendRow(headers); dbSheet.setFrozenRows(1); dbSheet.getRange("K:N").setNumberFormat("@"); 
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
        dateAdd: val(20), dateMod: val(21), lastSeen: parseDateLoose(val(22)) || new Date(0), lastCount: val(23), dateDel: val(25), tags: val(26), dateInv: val(27), qty: safeNumber(val(28)) 
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
      safeString(item.dateAdd), safeString(item.dateMod), formatDateForDisplay(item.lastSeen), 
      safeString(item.lastCount), "", safeString(item.dateDel), safeString(item.tags), safeString(item.dateInv), formatQty(item.qty) 
    ]); 
  }); 
  if (outputData.length > 0) { 
    if(dbSheet.getLastRow() > 1) dbSheet.getRange(2, 1, dbSheet.getLastRow()-1, dbSheet.getLastColumn()).clearContent(); 
    const CHUNK_SIZE = 1000; 
    for (let i = 0; i < outputData.length; i += CHUNK_SIZE) { 
      const chunk = outputData.slice(i, i + CHUNK_SIZE); 
      if(chunk.length > 0) { dbSheet.getRange(2 + i, 1, chunk.length, headers.length).setValues(chunk); SpreadsheetApp.flush(); } 
    } 
  } 
}

function deleteImportSession(sessionId, sheetName) { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const logSheet = ss.getSheetByName(LOG_SHEET_NAME); if(sheetName) { const sheet = ss.getSheetByName(sheetName); if(sheet) ss.deleteSheet(sheet); } const data = logSheet.getDataRange().getValues(); const idx = data.findIndex(r => String(r[0]) === String(sessionId)); if(idx > -1) { logSheet.deleteRow(idx + 1); return { success: true }; } return { success: false, error: "Session non trouvée dans les logs" }; }
function getImportHistory() { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const logSheet = ss.getSheetByName(LOG_SHEET_NAME); if(!logSheet) return []; const data = logSheet.getDataRange().getValues(); if (data.length <= 1) return []; const history = data.slice(1).map(r => { if (!r[0]) return null; const mixedName = String(r[3] || ""); const sheetName = mixedName.split(' | ')[0].trim(); const sheetExists = sheetName ? (ss.getSheetByName(sheetName) ? true : false) : false; return { id: String(r[0]), name: String(r[1]), timestamp: formatDateForDisplay(r[2]), sheetName: sheetName, fullLog: mixedName, exists: sheetExists }; }).filter(item => item !== null).reverse(); return history; }
function parseInventoryFromValues(values) { if (!values || values.length < 2) return { data: [], stockIdx: -1 }; const normalizeHeader = (h) => String(h).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(); const headers = values[0].map(normalizeHeader); const data = values.slice(1); const parseNum = (val) => { if (typeof val === 'string') { let clean = val.replace(/[\s\u00A0]/g, '').replace(',', '.'); return parseFloat(clean) || 0; } return (typeof val === 'number') ? val : 0; }; let stockIdx = headers.findIndex(h => h === 'quantite en stock'); if (stockIdx === -1) stockIdx = headers.findIndex(h => h === 'stock'); if (stockIdx === -1) { stockIdx = headers.findIndex(h => h.includes('stock') && !h.includes('sur') && !h.includes('valeur') && !h.includes('alert')); } const col = (k) => headers.findIndex(h => h.includes(k.toLowerCase())); const rackF_idx = col('rack f') > -1 ? col('rack f') : col('rack'); const colMap = { ref: col('ref'), stock: stockIdx, name: col('article/nom') > -1 ? col('article/nom') : col('nom'), suppName: col('fournisseurs/nom'), suppRef: col('fournisseurs/référence'), cote1: col('cote 1'), cote2: col('cote 2'), cote3: col('cote 3'), poids: col('poids'), rackF: rackF_idx, rowF: col('row f'), caseF: col('case f'), surF: col('surstock f'), rackG: col('rack g'), rowG: col('row g'), caseG: col('case g'), surG: col('surstock g'), rackA: col('rack a'), rowA: col('row a'), caseA: col('case a'), surA: col('surstock a'), tags: col('tags') > -1 ? col('tags') : col('etiquette') }; const parsedData = data.map(row => { const getVal = (idx) => (idx > -1 && row[idx] !== undefined) ? row[idx] : ""; const stockRaw = colMap.stock > -1 ? parseNum(row[colMap.stock]) : 0; return { ref: normalizeRef(String(row[colMap.ref])), name: colMap.name > -1 ? String(row[colMap.name]).trim() : "N/A", qty: stockRaw, suppName: String(getVal(colMap.suppName)).trim(), suppRef: String(getVal(colMap.suppRef)).trim(), rackF: safeVal(getVal(colMap.rackF)), rowF: safeVal(getVal(colMap.rowF)), caseF: safeVal(getVal(colMap.caseF)), surF: safeVal(getVal(colMap.surF)), rackG: safeVal(getVal(colMap.rackG)), rowG: safeVal(getVal(colMap.rowG)), caseG: safeVal(getVal(colMap.caseG)), surG: safeVal(getVal(colMap.surG)), rackA: safeVal(getVal(colMap.rackA)), rowA: safeVal(getVal(colMap.rowA)), caseA: safeVal(getVal(colMap.caseA)), surA: safeVal(getVal(colMap.surA)), cote1: colMap.cote1 > -1 ? parseNum(row[colMap.cote1]) : 0, cote2: colMap.cote2 > -1 ? parseNum(row[colMap.cote2]) : 0, cote3: colMap.cote3 > -1 ? parseNum(row[colMap.cote3]) : 0, poids: colMap.poids > -1 ? parseNum(row[colMap.poids]) : 0, tags: safeString(getVal(colMap.tags)) }; }).filter(item => item.ref !== ""); return { data: parsedData, stockIdx: stockIdx }; }
function saveImportToHistory(c, n, f) { try { const ss = SpreadsheetApp.openById(DB_SPREADSHEET_ID); const sep = detectSeparator(c); const csv = Utilities.parseCsv(c, sep); if(!csv||csv.length===0) throw new Error("Vide"); let sN = n.replace(/\//g, "-"); let s = ss.getSheetByName(sN); if(s) sN += "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmm"); s = ss.insertSheet(sN); if(csv.length){ s.getRange(1,1,csv.length,csv[0].length).setNumberFormat("@").setValues(csv); } let lS = ss.getSheetByName(LOG_SHEET_NAME); if(!lS) { lS=ss.insertSheet(LOG_SHEET_NAME); lS.appendRow(["ID","Nom","Date","Onglet"]); } const id=Date.now().toString(); const dStr=Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"); const res = parseInventoryFromValues(csv); lS.appendRow([id, n, dStr, sN + " | Stock Col: " + res.stockIdx]); updateMasterDbIncremental(ss, res, new Date(), n); return {success:true, message:"Sauvegardé", id:id}; } catch(e){ throw new Error(e.message); } }
function detectSeparator(csvString) { const firstLine = csvString.split('\n')[0]; return (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ','; }
function processInventoryInMemory(dbMap, newData, importDateObj, importDateStr) { const importDateStrFull = formatDateForDisplay(importDateObj); newData.forEach(item => { const ref = item.ref; const existing = dbMap.get(ref); if (existing) { if(item.name && item.name !== "N/A") existing.name = item.name; if(item.suppName) existing.suppName = item.suppName; if(item.suppRef) existing.suppRef = item.suppRef; if(item.rackF) existing.rackF = item.rackF; if(item.rowF) existing.rowF = item.rowF; if(item.caseF) existing.caseF = item.caseF; if(item.surF) existing.surF = item.surF; if(item.rackG) existing.rackG = item.rackG; if(item.rowG) existing.rowG = item.rowG; if(item.caseG) existing.caseG = item.caseG; if(item.surG) existing.surG = item.surG; if(item.rackA) existing.rackA = item.rackA; if(item.rowA) existing.rowA = item.rowA; if(item.caseA) existing.caseA = item.caseA; if(item.surA) existing.surA = item.surA; if(item.cote1) existing.cote1 = item.cote1; if(item.cote2) existing.cote2 = item.cote2; if(item.cote3) existing.cote3 = item.cote3; if(item.poids) existing.poids = item.poids; if(item.tags) existing.tags = item.tags; existing.qty = item.qty; existing.dateMod = importDateStrFull; existing.lastSeen = importDateObj; existing.dateInv = importDateStrFull; } else { dbMap.set(ref, { ref: ref, name: item.name || "Nouveau", suppName: item.suppName, suppRef: item.suppRef, rackF: item.rackF, rowF: item.rowF, caseF: item.caseF, surF: item.surF, rackG: item.rackG, rowG: item.rowG, caseG: item.caseG, surG: item.surG, rackA: item.rackA, rowA: item.rowA, caseA: item.caseA, surA: item.surA, cote1: item.cote1, cote2: item.cote2, cote3: item.cote3, poids: item.poids, dateAdd: importDateStrFull, dateMod: importDateStrFull, lastSeen: importDateObj, lastCount: "", dateDel: "", tags: item.tags, dateInv: importDateStrFull, qty: item.qty }); } }); }
