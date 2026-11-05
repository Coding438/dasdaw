/**
 * Forge Studio — Client Portal
 * Google Apps Script backend
 *
 * Web App URL (your deployment):
 * https://script.google.com/macros/s/AKfycbx3ZgE_IEf2kX2KIliRdPf7K41GiGzXNNjs4IRC2Zcis4g4qdDKxWQDixdZ328s1oPfZg/exec
 *
 * SETUP:
 * 1. Create a new Google Spreadsheet
 * 2. Extensions → Apps Script
 * 3. Paste this Code.gs
 * 4. Create a file named Index.html and paste the HTML
 * 5. Run setupSheets() once (from the editor)
 * 6. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (or your org)
 */

const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbx3ZgE_IEf2kX2KIliRdPf7K41GiGzXNNjs4IRC2Zcis4g4qdDKxWQDixdZ328s1oPfZg/exec';

const SHEET_NAMES = {
  clients: 'Clients',
  projects: 'Projects',
  leads: 'Leads',
  invoices: 'Invoices',
  team: 'Team',
  activity: 'Activity',
  settings: 'Settings'
};

/* ============================================================
   SETUP — creates all sheets + headers + sample data
   ============================================================ */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Clients ----
  let sh = getOrCreateSheet_(ss, SHEET_NAMES.clients, [
    'id', 'company', 'initials', 'contact', 'email', 'phone',
    'industry', 'status', 'projects', 'billed', 'since'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 6, 11).setValues([
      ['CL-101', 'Northwind Retail', 'NR', 'Priya Raman', 'priya@northwind.com', '+92 (42) 111-0142', 'Retail', 'Active', 1, 9720000, '2023'],
      ['CL-102', 'Vertex Labs', 'VL', 'Sam Ortega', 'sam@vertexlabs.io', '+92 (21) 111-2233', 'SaaS', 'Active', 1, 8380000, '2024'],
      ['CL-103', 'Harbor Health', 'HH', 'Leila Hassan', 'leila@harbor.health', '+92 (51) 111-4455', 'Healthcare', 'Active', 1, 6500000, '2024'],
      ['CL-104', 'Lumen Co.', 'LC', 'Chris Vale', 'chris@lumenco.co', '+92 (42) 111-6677', 'Design', 'Prospect', 1, 1800000, '2025'],
      ['CL-105', 'Pinecrest Hotels', 'PH', 'Nadia Khan', 'nadia@pinecrest.pk', '+92 (21) 111-8899', 'Hospitality', 'Active', 1, 590000, '2025'],
      ['CL-106', 'Atlas Freight', 'AF', 'Omar Farid', 'omar@atlasfreight.com', '+92 (42) 111-0011', 'Logistics', 'Churned', 1, 11360000, '2022']
    ]);
  }

  // ---- Projects ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.projects, [
    'id', 'name', 'client', 'status', 'progress', 'due',
    'team', 'tags', 'budget', 'spent'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 6, 10).setValues([
      ['PRJ-2041', 'Northwind Commerce Replatform', 'Northwind Retail', 'In progress', 72, 'Mar 14', 'MK,JD,SP', 'Next.js,Shopify,Figma', 13500000, 9720000],
      ['PRJ-2038', 'Vertex Analytics Dashboard', 'Vertex Labs', 'In review', 91, 'Feb 28', 'AL,RB', 'React,D3.js,Node', 9000000, 8380000],
      ['PRJ-2035', 'Harbor Health Booking Flow', 'Harbor Health', 'In progress', 45, 'Apr 02', 'JD,TN,MK', 'Next.js,Stripe', 15750000, 6500000],
      ['PRJ-2029', 'Lumen Brand Site', 'Lumen Co.', 'Blocked', 28, 'Apr 18', 'SP', 'Astro,GSAP', 5200000, 1800000],
      ['PRJ-2044', 'Pinecrest CMS Migration', 'Pinecrest Hotels', 'In progress', 12, 'May 09', 'TN,RB', 'Sanity,Next.js', 7600000, 590000],
      ['PRJ-2022', 'Atlas Mobile Web App', 'Atlas Freight', 'Completed', 100, 'Delivered', 'RB,AL', 'PWA,Firebase', 11500000, 11360000]
    ]);
  }

  // ---- Leads ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.leads, [
    'id', 'company', 'contact', 'email', 'phone', 'value',
    'stage', 'source', 'next'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 5, 9).setValues([
      ['LD-301', 'Brightline Media', 'Dana Cole', 'dana@brightline.io', '+92 (300) 123-4567', 6720000, 'Proposal', 'Referral', 'Mar 02'],
      ['LD-302', 'Cobalt Fitness', 'Ravi Shah', 'ravi@cobaltfit.com', '+92 (321) 987-6543', 3360000, 'Qualified', 'Website', 'Feb 27'],
      ['LD-303', 'Nova Fintech', 'Iris Wald', 'iris@novafin.com', '+92 (333) 456-7890', 15680000, 'Contacted', 'Outbound', 'Mar 05'],
      ['LD-304', 'Greenline Energy', 'Omar Haddad', 'omar@greenline.energy', '+92 (345) 234-5678', 10640000, 'New', 'LinkedIn', 'Feb 26'],
      ['LD-305', 'Studio Kite', 'June Park', 'june@studiokite.design', '+92 (312) 876-5432', 2660000, 'New', 'Website', 'Mar 01']
    ]);
  }

  // ---- Invoices ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.invoices, [
    'id', 'client', 'project', 'amount', 'issued', 'due', 'status'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 6, 7).setValues([
      ['INV-1046', 'Northwind Retail', 'Commerce Replatform', 3360000, 'Feb 01, 2025', 'Feb 15, 2025', 'Overdue'],
      ['INV-1045', 'Northwind Retail', 'Commerce Replatform', 2800000, 'Jan 15, 2025', 'Jan 29, 2025', 'Paid'],
      ['INV-1044', 'Vertex Labs', 'Analytics Dashboard', 2240000, 'Jan 28, 2025', 'Feb 11, 2025', 'Paid'],
      ['INV-1043', 'Harbor Health', 'Booking Flow', 2772000, 'Jan 20, 2025', 'Feb 03, 2025', 'Paid'],
      ['INV-1042', 'Lumen Co.', 'Brand Site', 1736000, 'Jan 12, 2025', 'Jan 26, 2025', 'Overdue'],
      ['INV-1041', 'Pinecrest Hotels', 'CMS Migration', 980000, 'Feb 05, 2025', 'Feb 19, 2025', 'Pending']
    ]);
  }

  // ---- Team ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.team, [
    'name', 'initials', 'role', 'projects', 'capacity', 'skills'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 6, 6).setValues([
      ['Maya Kim', 'MK', 'Design Lead', 3, 72, 'Figma,Design systems,Motion'],
      ['Jonas Dale', 'JD', 'Senior Engineer', 2, 95, 'Next.js,TypeScript,Node'],
      ['Sofia Petrov', 'SP', 'Frontend Engineer', 2, 48, 'React,GSAP,Tailwind'],
      ['Ravi Bhatt', 'RB', 'Backend Engineer', 2, 88, 'Node,Postgres,AWS'],
      ['Alina Lopez', 'AL', 'Product Strategist', 3, 60, 'Discovery,Analytics,Roadmaps'],
      ['Tomás Nunes', 'TN', 'QA & Delivery', 2, 35, 'Playwright,Cypress,CI/CD']
    ]);
  }

  // ---- Activity ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.activity, [
    'icon', 'text', 'time'
  ]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 5, 3).setValues([
      ['code', '<b>Maya Kim</b> pushed 12 commits to <b>Northwind Replatform</b>', '2 hours ago'],
      ['money', '<b>Vertex Labs</b> paid invoice <b>INV-1044</b> — Rs 2,240,000', '5 hours ago'],
      ['move', '<b>Jonas Dale</b> moved <b>Harbor Health</b> to review', 'Yesterday, 4:12 PM'],
      ['check', '<b>Atlas Mobile Web App</b> was marked complete', 'Yesterday, 11:03 AM'],
      ['upload', '<b>Sofia Petrov</b> uploaded 6 assets to <b>Lumen Brand Site</b>', '2 days ago']
    ]);
  }

  // ---- Settings ----
  sh = getOrCreateSheet_(ss, SHEET_NAMES.settings, ['key', 'value']);
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, 4, 2).setValues([
      ['currency', 'PKR'],
      ['studioName', 'Forge Studio'],
      ['timezone', 'UTC+05:00'],
      ['theme', 'dark']
    ]);
  }

  SpreadsheetApp.flush();
  return 'All sheets created / verified successfully.';
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  // Always ensure headers are present in row 1
  const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = existing.every(c => c === '' || c === null);
  if (needsHeaders || sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1c2029')
      .setFontColor('#e9ecf3');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ============================================================
   WEB APP ENTRY
   ============================================================ */
function doGet(e) {
  // Ensure sheets exist on every open (safe & idempotent)
  try { setupSheets(); } catch (err) { Logger.log(err); }

  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';

  // JSON API mode (for external HTML / fetch)
  if (action === 'getAllData') {
    return jsonOut_(getAllData());
  }
  if (action === 'ping') {
    return jsonOut_({ ok: true, url: WEB_APP_URL });
  }

  // Default: serve the portal HTML
  const template = HtmlService.createTemplateFromFile('Index');
  template.WEB_APP_URL = WEB_APP_URL;
  return template.evaluate()
    .setTitle('Forge Studio — Client Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * POST body JSON: { action, record, id, stage, ... }
 * Supports: saveProject, saveClient, saveLead, deleteLead,
 *           deleteClient, deleteProject, updateLeadStage
 */
function doPost(e) {
  try { setupSheets(); } catch (err) { Logger.log(err); }

  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return jsonOut_({ error: 'Invalid JSON body' });
  }

  const action = body.action || '';
  try {
    switch (action) {
      case 'saveProject':
        return jsonOut_(saveProject(body.record || {}));
      case 'saveClient':
        return jsonOut_(saveClient(body.record || {}));
      case 'saveLead':
        return jsonOut_(saveLead(body.record || {}));
      case 'saveInvoice':
        return jsonOut_(saveInvoice(body.record || {}));
      case 'deleteProject':
        return jsonOut_(deleteProject(body.id));
      case 'deleteClient':
        return jsonOut_(deleteClient(body.id));
      case 'deleteLead':
        return jsonOut_(deleteLead(body.id));
      case 'updateLeadStage':
        return jsonOut_(updateLeadStage(body.id, body.stage));
      case 'getAllData':
        return jsonOut_(getAllData());
      default:
        return jsonOut_({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut_({ error: String(err.message || err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
   DATA API  (called via google.script.run)
   ============================================================ */
function getAllData() {
  setupSheets(); // safety
  return {
    clients: sheetToObjects_(SHEET_NAMES.clients),
    projects: sheetToObjects_(SHEET_NAMES.projects).map(normalizeProject_),
    leads: sheetToObjects_(SHEET_NAMES.leads),
    invoices: sheetToObjects_(SHEET_NAMES.invoices),
    team: sheetToObjects_(SHEET_NAMES.team).map(normalizeTeam_),
    activity: sheetToObjects_(SHEET_NAMES.activity)
  };
}

function normalizeProject_(p) {
  return {
    ...p,
    progress: Number(p.progress) || 0,
    budget: Number(p.budget) || 0,
    spent: Number(p.spent) || 0,
    team: (p.team || '').toString().split(',').map(s => s.trim()).filter(Boolean),
    tags: (p.tags || '').toString().split(',').map(s => s.trim()).filter(Boolean)
  };
}

function normalizeTeam_(t) {
  return {
    ...t,
    projects: Number(t.projects) || 0,
    capacity: Number(t.capacity) || 0,
    skills: (t.skills || '').toString().split(',').map(s => s.trim()).filter(Boolean)
  };
}

function sheetToObjects_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .filter(row => row.some(c => c !== '' && c !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

/* ---------- Save / Update helpers ---------- */
function saveRow_(sheetName, idField, record) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf(idField);

  let rowIndex = -1;
  if (record[idField]) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(record[idField])) {
        rowIndex = i + 1;
        break;
      }
    }
  }

  const row = headers.map(h => {
    let v = record[h];
    if (Array.isArray(v)) v = v.join(',');
    return v === undefined || v === null ? '' : v;
  });

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return true;
}

function deleteById_(sheetName, idField, id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf(idField);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === String(id)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/* ---------- Public CRUD ---------- */
function saveProject(record) {
  if (!record.id) record.id = 'PRJ-' + (2000 + Math.floor(Math.random() * 900));
  if (Array.isArray(record.team)) record.team = record.team.join(',');
  if (Array.isArray(record.tags)) record.tags = record.tags.join(',');
  saveRow_(SHEET_NAMES.projects, 'id', record);
  return getAllData();
}

function deleteProject(id) {
  deleteById_(SHEET_NAMES.projects, 'id', id);
  return getAllData();
}

function saveClient(record) {
  if (!record.id) record.id = 'CL-' + (100 + Math.floor(Math.random() * 900));
  if (!record.initials && record.company) {
    record.initials = record.company.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  }
  saveRow_(SHEET_NAMES.clients, 'id', record);
  return getAllData();
}

function deleteClient(id) {
  deleteById_(SHEET_NAMES.clients, 'id', id);
  return getAllData();
}

function saveLead(record) {
  if (!record.id) record.id = 'LD-' + (300 + Math.floor(Math.random() * 700));
  saveRow_(SHEET_NAMES.leads, 'id', record);
  return getAllData();
}

function deleteLead(id) {
  deleteById_(SHEET_NAMES.leads, 'id', id);
  return getAllData();
}

function updateLeadStage(id, stage) {
  const leads = sheetToObjects_(SHEET_NAMES.leads);
  const lead = leads.find(l => String(l.id) === String(id));
  if (!lead) throw new Error('Lead not found');
  lead.stage = stage;
  saveRow_(SHEET_NAMES.leads, 'id', lead);
  return getAllData();
}

function saveInvoice(record) {
  if (!record.id) record.id = 'INV-' + (1000 + Math.floor(Math.random() * 900));
  saveRow_(SHEET_NAMES.invoices, 'id', record);
  return getAllData();
}

function logActivity(icon, text) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.activity);
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 3).setValues([[icon || 'check', text, 'Just now']]);
  return true;
}
