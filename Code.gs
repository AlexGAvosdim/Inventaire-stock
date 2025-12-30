/*** ==============================================================================
* PROJET : GESTIONNAIRE SUPPLY CHAIN - AVOSDIM
* VERSION : 7.1.0 (ARCHITECTURE MODULAIRE 8 FICHIERS)
* DATE : 30/12/2025
* ==============================================================================
*/

// --- CONFIGURATION ---
const DB_SPREADSHEET_ID = "1Gz-xX5YCaMQkaQBX83SJzKQ-l_D8_ARjEHu9ynPwcxo";
const LOG_SHEET_NAME = "LOG";
const MASTER_DB_SHEET_NAME = "BASE_ARTICLES";

// --- MOTEUR DE TEMPLATE ---
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() { 
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Avosdim Supply Chain v7.1.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

// --- MENU ADMIN ---
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('AVOSDIM Admin')
      .addItem('▶️ Autoriser l\'automatisation', 'debugTriggerPermission')
      .addToUi();
  } catch (e) {}
}

function debugTriggerPermission() {
  ScriptApp.getProjectTriggers();
  try { SpreadsheetApp.getUi().alert("✅ Autorisation validée !"); } catch(e) {}
}

// --- UTILS PARTAGÉS ---
function safeVal(v) { if (v === 0 || v === "0") return "0"; if (!v) return ""; return String(v); }
function safeString(v) { return (v === null || v === undefined) ? "" : String(v); }
function safeNumber(v) { return (v === null || v === undefined || isNaN(v) || v === "") ? 0 : Number(v); }
function cleanSurstock(v) { if (!v) return ""; return String(v).trim().replace(/\//g, " | "); }
function formatDateForDisplay(val) { if (!val) return ""; if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"); return String(val); }
function parseDateLoose(val) { if (!val) return null; if (val instanceof Date) return val; try { if (typeof val === 'string' && val.includes('/')) { const p = val.split('/'); if (p.length === 3) return new Date(p[2], p[1]-1, p[0]); } const d = new Date(val); return isNaN(d.getTime()) ? null : d; } catch(e) { return null; } }
function normalizeRef(v) { if (v === null || v === undefined) return ""; let s = String(v).trim(); if (/^\d+$/.test(s)) { return s.replace(/^0+/, '') || "0"; } return s; }
function formatQty(val) { const n = Number(val); if (isNaN(n)) return 0; return Math.round(n * 1000) / 1000; }
function ensureSheetDimensions(sheet, requiredCols) { const cur = sheet.getMaxColumns(); if (cur < requiredCols) sheet.insertColumnsAfter(cur, requiredCols - cur); }
