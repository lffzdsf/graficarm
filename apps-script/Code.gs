const SETTINGS = Object.freeze({
  spreadsheetId: '1bDb2u_s19E2JMJd73J_ovP-l59bs5LCRkdNM2YgpJ6c',
  workflowFolderId: '1Z6jiTVZ6UEO6FwlykwjuC4ruQv4icaw-',
  demandSheet: 'Demandas',
  technicalSheet: 'Controle Técnico',
  logSheet: 'Logs do Sistema',
  maxFileBytes: 8 * 1024 * 1024,
  allowedMimeTypes: ['image/jpeg', 'image/png', 'image/webp', 'application/pdf']
});

function doGet() {
  return json_({ok: true, service: 'Graficarm API', version: 2});
}

function doPost(e) {
  try {
    const action = String(e.parameter.action || 'list');
    const payload = JSON.parse(e.parameter.payload || '{}');
    if (action === 'list') return json_({ok: true, demands: listDemands_()});
    if (action === 'file') return json_(filePreview_(payload));
    if (action === 'save') return json_(withLock_(() => saveDemand_(payload)));
    if (action === 'delete') return json_(withLock_(() => deleteDemand_(payload)));
    throw new Error('Ação inválida.');
  } catch (error) {
    return json_({ok: false, error: error.message || String(error)});
  }
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return callback(); } finally { lock.releaseLock(); }
}

function listDemands_() {
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const sheet = ss.getSheetByName(SETTINGS.demandSheet);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rows = sheet.getRange(2, 1, lastRow - 1, 22).getDisplayValues();
  const checks = technicalChecks_(ss);
  return rows.filter(row => row[1]).map(row => {
    const link = row[16];
    return {
      id: row[0], name: row[1], category: row[2], city: row[3], format: row[4],
      priority: row[5], status: row[6], owner: row[7], requester: row[8],
      createdAt: brDateToIso_(row[9]), deadline: brDateToIso_(row[10]),
      printDate: brDateToIso_(row[11]), quantity: row[13], supplier: row[14],
      version: row[15], link: link, attachment: attachmentInfo_(link),
      adjustments: row[17], approvedBy: row[18],
      approvalDate: brDateToIso_(row[19]), updatedAt: brDateToIso_(row[20]),
      notes: row[21], checks: checks[row[0]] || emptyChecks_()
    };
  });
}

function saveDemand_(payload) {
  const incoming = payload.demand || {};
  if (!String(incoming.name || '').trim()) throw new Error('Informe o nome do material.');
  if (!String(incoming.requester || '').trim()) throw new Error('Informe o solicitante.');
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const sheet = ss.getSheetByName(SETTINGS.demandSheet);
  let rowNumber = findDemandRow_(sheet, incoming.id);
  if (!rowNumber) rowNumber = firstAvailableDemandRow_(sheet);
  const current = sheet.getRange(rowNumber, 1, 1, 22).getValues()[0];
  const id = current[0] || nextId_(sheet);
  let link = String(incoming.link || current[16] || '').trim();
  if (payload.file) link = uploadFile_(id, payload.file);
  const now = new Date();
  const firstBlock = [[
    id,
    clean_(incoming.name, 180),
    clean_(incoming.category || 'Outros', 80),
    clean_(incoming.city || 'Todo o estado', 100),
    clean_(incoming.format, 100),
    clean_(incoming.priority || 'Média', 20),
    clean_(incoming.status || 'Solicitação', 40),
    clean_(incoming.owner || 'A definir', 100),
    clean_(incoming.requester, 100),
    current[9] || isoToDate_(incoming.createdAt) || now,
    isoToDate_(incoming.deadline),
    isoToDate_(incoming.printDate)
  ]];
  sheet.getRange(rowNumber, 1, 1, 12).setValues(firstBlock);
  sheet.getRange(rowNumber, 14, 1, 9).setValues([[
    clean_(incoming.quantity, 80),
    clean_(incoming.supplier, 120),
    clean_(incoming.version || current[15] || 'v01', 30),
    link,
    clean_(incoming.adjustments, 1000),
    clean_(incoming.approvedBy, 100),
    isoToDate_(incoming.approvalDate),
    now,
    clean_(incoming.notes, 2000)
  ]]);
  upsertTechnical_(ss, id, incoming.checks || emptyChecks_(), incoming.owner || 'A definir');
  log_(ss, 'SALVAR', id, incoming.requester, incoming.status || 'Solicitação');
  SpreadsheetApp.flush();
  const saved = listDemands_().find(item => item.id === id);
  return {ok: true, demand: saved, fileUrl: link};
}

function deleteDemand_(payload) {
  const id = clean_(payload.id, 30);
  if (!id) throw new Error('Informe a demanda que será removida.');
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const sheet = ss.getSheetByName(SETTINGS.demandSheet);
  const rowNumber = findDemandRow_(sheet, id);
  if (!rowNumber) throw new Error('A demanda não foi encontrada ou já foi removida.');
  const row = sheet.getRange(rowNumber, 1, 1, 22).getDisplayValues()[0];
  const actor = row[8] || row[7] || 'Interface pública';
  log_(ss, 'REMOVER', id, actor, 'Demanda removida; arquivos do Drive preservados.');
  sheet.getRange(rowNumber, 1, 1, 12).clearContent();
  sheet.getRange(rowNumber, 14, 1, 9).clearContent();
  const technical = ss.getSheetByName(SETTINGS.technicalSheet);
  const technicalRow = findRowByValue_(technical, 1, id);
  if (technicalRow) {
    technical.getRange(technicalRow, 1).clearContent();
    technical.getRange(technicalRow, 3, 1, 8).clearContent();
    technical.getRange(technicalRow, 12, 1, 2).clearContent();
  }
  SpreadsheetApp.flush();
  return {ok: true, deletedId: id, filesPreserved: true};
}

function attachmentInfo_(link) {
  if (!link) return null;
  const match = String(link).match(/[-\w]{25,}/);
  if (!match) return {name: 'Arquivo da demanda', mime: '', url: String(link), previewUrl: String(link), thumbnailUrl: ''};
  try {
    const file = DriveApp.getFileById(match[0]);
    const mime = file.getMimeType();
    return {
      id: file.getId(),
      name: file.getName(),
      mime: mime,
      url: file.getUrl(),
      previewUrl: 'https://drive.google.com/file/d/' + file.getId() + '/preview',
      thumbnailUrl: mime.indexOf('image/') === 0 ? 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200' : ''
    };
  } catch (error) {
    return {name: 'Arquivo da demanda', mime: '', url: String(link), previewUrl: String(link), thumbnailUrl: ''};
  }
}

function filePreview_(payload) {
  const id = clean_(payload.id, 30);
  if (!id) throw new Error('Informe a demanda do arquivo.');
  const ss = SpreadsheetApp.openById(SETTINGS.spreadsheetId);
  const sheet = ss.getSheetByName(SETTINGS.demandSheet);
  const rowNumber = findDemandRow_(sheet, id);
  if (!rowNumber) throw new Error('Demanda não encontrada.');
  const link = sheet.getRange(rowNumber, 17).getDisplayValue();
  const match = String(link).match(/[-\w]{25,}/);
  if (!match) throw new Error('Esta demanda não possui um arquivo do Drive para visualizar.');
  const file = DriveApp.getFileById(match[0]);
  const mime = file.getMimeType();
  if (!SETTINGS.allowedMimeTypes.includes(mime)) throw new Error('O arquivo não possui um formato de pré-visualização permitido.');
  if (file.getSize() > SETTINGS.maxFileBytes) throw new Error('O arquivo excede o limite de pré-visualização de 8 MB.');
  return {
    ok: true,
    name: file.getName(),
    mime: mime,
    url: file.getUrl(),
    base64: Utilities.base64Encode(file.getBlob().getBytes())
  };
}

function technicalChecks_(ss) {
  const sheet = ss.getSheetByName(SETTINGS.technicalSheet);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const map = {};
  rows.forEach(row => {
    if (!row[0]) return;
    map[row[0]] = {cmyk: !!row[2], dpi: !!row[3], bleed: !!row[4], margin: !!row[5], fonts: !!row[6], legal: !!row[7], pdf: !!row[8], proof: !!row[9]};
  });
  return map;
}

function upsertTechnical_(ss, id, checks, owner) {
  const sheet = ss.getSheetByName(SETTINGS.technicalSheet);
  let row = findRowByValue_(sheet, 1, id);
  if (!row) row = firstBlankRow_(sheet, 1);
  sheet.getRange(row, 1).setValue(id);
  sheet.getRange(row, 3, 1, 8).setValues([[
    !!checks.cmyk, !!checks.dpi, !!checks.bleed, !!checks.margin,
    !!checks.fonts, !!checks.legal, !!checks.pdf, !!checks.proof
  ]]);
  sheet.getRange(row, 12).setValue(clean_(owner, 100));
  sheet.getRange(row, 13).setValue(new Date());
}

function uploadFile_(id, file) {
  const mime = String(file.mime || '');
  const bytes = Utilities.base64Decode(String(file.base64 || ''));
  if (!SETTINGS.allowedMimeTypes.includes(mime)) throw new Error('Tipo de arquivo não permitido.');
  if (!bytes.length || bytes.length > SETTINGS.maxFileBytes) throw new Error('O arquivo deve ter no máximo 8 MB.');
  const root = DriveApp.getFolderById(SETTINGS.workflowFolderId);
  const filesFolder = getOrCreateFolder_(root, 'Arquivos da Central');
  const demandFolder = getOrCreateFolder_(filesFolder, id);
  const safeName = String(file.name || 'arquivo').replace(/[^\w.() áàâãéêíóôõúç-]/gi, '_').slice(0, 140);
  const created = demandFolder.createFile(Utilities.newBlob(bytes, mime, safeName));
  return created.getUrl();
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function findDemandRow_(sheet, id) {
  if (!id) return 0;
  return findRowByValue_(sheet, 1, id);
}

function findRowByValue_(sheet, column, value) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const finder = sheet.getRange(2, column, lastRow - 1, 1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : 0;
}

function firstAvailableDemandRow_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const values = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  const index = values.findIndex(row => !row[0]);
  if (index >= 0) return index + 2;
  sheet.insertRowAfter(lastRow);
  return lastRow + 1;
}

function firstBlankRow_(sheet, column) {
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const values = sheet.getRange(2, column, lastRow - 1, 1).getDisplayValues();
  const index = values.findIndex(row => !row[0]);
  if (index >= 0) return index + 2;
  sheet.insertRowAfter(lastRow);
  return lastRow + 1;
}

function nextId_(sheet) {
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getDisplayValues().flat();
  const max = ids.reduce((acc, id) => Math.max(acc, Number(String(id).replace(/\D/g, '')) || 0), 0);
  return 'RM-' + String(max + 1).padStart(4, '0');
}

function log_(ss, action, id, actor, detail) {
  let sheet = ss.getSheetByName(SETTINGS.logSheet);
  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS.logSheet);
    sheet.getRange(1, 1, 1, 5).setValues([['Data / Hora', 'Ação', 'ID', 'Responsável informado', 'Detalhe']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), action, id, clean_(actor, 100), clean_(detail, 500)]);
}

function emptyChecks_() { return {cmyk:false,dpi:false,bleed:false,margin:false,fonts:false,legal:false,pdf:false,proof:false}; }
function clean_(value, limit) { return String(value == null ? '' : value).replace(/[<>]/g, '').trim().slice(0, limit || 500); }
function isoToDate_(value) { if (!value) return ''; const d = new Date(String(value).slice(0,10) + 'T12:00:00'); return isNaN(d) ? '' : d; }
function brDateToIso_(value) { const match = String(value || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/); return match ? `${match[3]}-${match[2]}-${match[1]}` : ''; }
function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
