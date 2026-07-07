/**
 * Bulletin Board — multi-tenant Apps Script backend
 * One script serves all buildings.
 *
 * Master spreadsheet contains a 'clients' sheet:
 *   A: client id (e.g. tamar)  B: spreadsheet ID  C: active (כן/לא)
 *
 * Each building spreadsheet contains sheets:
 *   settings — A: key (Hebrew), B: value
 *   notices  — A: icon, B: title, C: text
 *   central  — A: image URL (optional, overrides Drive folder)
 *   footer   — A: message
 */

const MASTER_SPREADSHEET_ID = 'PASTE_MASTER_SPREADSHEET_ID_HERE';
const IMAGES_CACHE_SECONDS = 300;

function doGet(e) {
  try {
    const client = String((e && e.parameter && e.parameter.client) || '').trim();
    if (!client) return json({ error: 'missing ?client= parameter' });

    const targetId = lookupClient(client);
    if (!targetId) return json({ error: 'unknown or inactive client: ' + client });

    const ss = SpreadsheetApp.openById(targetId);

    const settings = readKeyValue(ss, 'settings');
    const notices = readRows(ss, 'notices')
      .map(function (r) { return { icon: str(r[0]), title: str(r[1]), text: str(r[2]) }; })
      .filter(function (n) { return n.title || n.text; });
    const central = readRows(ss, 'central')
      .map(function (r) { return str(r[0]); })
      .filter(Boolean);
    const footer = readRows(ss, 'footer')
      .map(function (r) { return str(r[0]); })
      .filter(Boolean);
    const images = getDriveImages(settings['תיקיית_תמונות']);

    return json({
      settings: settings,
      notices: notices,
      central: central,
      footer: footer,
      images: images
    });
  } catch (err) {
    return json({ error: String(err) });
  }
}

/* ---------- registry ---------- */
function lookupClient(client) {
  const master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const sheet = master.getSheetByName('clients');
  if (!sheet) throw new Error("master 'clients' sheet not found");
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    const id = str(rows[i][0]);
    const active = str(rows[i][2]);
    if (id === client && active !== 'לא') return str(rows[i][1]);
  }
  return null;
}

/* ---------- generic sheet readers ---------- */
function readKeyValue(ss, name) {
  const sheet = ss.getSheetByName(name);
  const out = {};
  if (!sheet) return out;
  const rows = sheet.getDataRange().getValues();
  for (var i = 0; i < rows.length; i++) {
    const key = str(rows[i][0]);
    if (key && key !== 'מפתח') out[key] = str(rows[i][1]); // skip header row
  }
  return out;
}

function readRows(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1).filter(function (r) {
    return r.join('').trim() !== '';
  });
}

/* ---------- Drive images as lightweight thumbnail URLs ---------- */
function getDriveImages(folderId) {
  folderId = str(folderId);
  if (!folderId) return [];

  const cache = CacheService.getScriptCache();
  const cacheKey = 'imgs_' + folderId;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const images = [];
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType().indexOf('image/') === 0) {
        images.push({
          // <img> tags are not subject to CORS — thumbnails load directly
          url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1920',
          name: file.getName()
        });
      }
    }
    // order controlled by filename: 01_name, 02_name...
    images.sort(function (a, b) { return a.name.localeCompare(b.name); });
    cache.put(cacheKey, JSON.stringify(images), IMAGES_CACHE_SECONDS);
    return images;
  } catch (err) {
    console.error('Drive folder error:', err);
    return [];
  }
}

/* ---------- utils ---------- */
function str(v) { return v === null || v === undefined ? '' : String(v).trim(); }

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
