/**
 * Social discernment study, two-question format (Prototype V13)
 *
 * Writes ONE ROW PER RESPONSE to the "Two-question responses" tab:
 *   Submitted at | Code | Name | Scenario | Time spent | Time spent (seconds) |
 *   Q1: prediction | Q1: voice note | Q2: number of highlights |
 *   [what changed after the participant first reached Review and submit: 12 columns, see REVIEW_HEADERS] |
 *   Q2 highlight 1: highlighted text | Q2 highlight 1: what it tells them | Q2 highlight 1: voice note |
 *   Q2 highlight 2: ... (a new block of three columns is added whenever someone makes more highlights)
 *
 * Also writes the same highlights to a "Q2 highlights" tab, one highlight per row, for analysis,
 * with a "Change in review" column. Highlights removed during review are listed there too.
 *
 * If the "Two-question responses" tab already exists from the earlier script, the review
 * columns are inserted into it automatically; older rows just have them blank.
 * Voice notes are saved as audio files in a private Drive folder and linked from the row.
 *
 * SETUP
 * 1. Open the Google Sheet the page already saves to: Extensions > Apps Script.
 * 2. Replace ALL the existing code with this file and click Save.
 * 3. Deploy > Manage deployments > pencil icon > Version: "New version" > Deploy.
 *    (Editing the code alone does NOT update the live web app. The URL stays the same.)
 * 4. Check: open the web app URL in an incognito window. It should say "two-question format, with review changes".
 */

const MAIN_TAB = 'Two-question responses';
const LONG_TAB = 'Q2 highlights';
const VOICE_FOLDER_NAME = 'Scenario study voice notes';

const CORE_HEADERS = [
  'Submitted at', 'Code', 'Name', 'Scenario', 'Time spent', 'Time spent (seconds)',
  'Q1: prediction', 'Q1: voice note', 'Q2: number of highlights'
];
const REVIEW_HEADERS = [
  'Time from first review to submit', 'Time from first review to submit (seconds)',
  'Q1 changed in review', 'Q1 words added / removed in review', 'Q1 prediction (before review)',
  'Highlights before review', 'Highlights added in review', 'Highlights removed in review',
  'Highlights reworded in review', 'Highlights re-highlighted in review',
  'Voice notes added / deleted in review', 'Q2 highlights (before review)'
];
const BASE_HEADERS = CORE_HEADERS.concat(REVIEW_HEADERS);
const BLOCK = ['highlighted text', 'what it tells them', 'voice note'];
const START_BLOCKS = 3;   // columns for 3 highlights are created up front; more are added as needed

const LONG_HEADERS = [
  'Submitted at', 'Code', 'Name', 'Scenario', 'Highlight #', 'Highlighted text', 'What it tells them', 'Voice note',
  'Change in review'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.format !== 'two-question') {
      return json_({ ok: false, error: 'This script expects the two-question page (V13 or later).' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const main = getSheet_(ss, MAIN_TAB, BASE_HEADERS.concat(blockHeaders_(1, START_BLOCKS)));
    const long = getSheet_(ss, LONG_TAB, LONG_HEADERS);
    upgradeMain_(main);
    upgradeLong_(long);

    // A retry of the same submission (same code) is ignored rather than written twice.
    if (main.getLastRow() > 1) {
      const codes = main.getRange(2, 2, main.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
      if (codes.indexOf(String(data.code)) !== -1) return json_({ ok: true, duplicate: true });
    }

    const highlights = data.highlights || [];
    ensureBlocks_(main, highlights.length);

    const links = saveVoiceNotes_(data);
    const voiceCell = v => !v ? '' : v.simulated ? 'Simulated (no audio)' : (links[v.id] || '');
    const submitted = new Date(data.submittedAt || Date.now());
    const secs = Number(data.secondsSpent) || 0;
    const p = data.prediction || {};

    const row = [
      submitted, data.code, data.name, data.scenario,
      duration_(secs), secs,
      p.text || '', voiceCell(p.voice), highlights.length
    ].concat(reviewCells_(data.review));
    highlights.forEach(h => row.push(h.quote || '', h.explanation || '', voiceCell(h.voice)));
    main.appendRow(row.map(safe_));

    const longRows = highlights.map((h, i) => [
      submitted, data.code, data.name, data.scenario, i + 1, h.quote || '', h.explanation || '', voiceCell(h.voice),
      h.reviewStatus || ''
    ].map(safe_));
    ((data.review && data.review.removedHighlights) || []).forEach(h => {
      longRows.push([
        submitted, data.code, data.name, data.scenario, '', h.quote || '', h.explanation || '',
        h.voice ? '(voice note deleted)' : '', 'Removed in review'
      ].map(safe_));
    });
    if (longRows.length) long.getRange(long.getLastRow() + 1, 1, longRows.length, LONG_HEADERS.length).setValues(longRows);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, message: 'Scenario study endpoint is running (two-question format, with review changes).' });
}

function duration_(secs) {
  secs = Number(secs) || 0;
  return Math.floor(secs / 60) + ' min ' + (secs % 60) + ' s';
}

// The 12 "what changed in review" cells, in the same order as REVIEW_HEADERS.
function reviewCells_(r) {
  if (!r) return REVIEW_HEADERS.map(() => '');
  const before = r.before || { prediction: {}, highlights: [] };
  const beforeList = (before.highlights || []).map((h, i) =>
    (i + 1) + '. "' + (h.quote || '') + '" \u2192 ' + (h.explanation || '(no text)') + (h.voice ? ' [voice note]' : '')
  ).join('\n');
  const bp = before.prediction || {};
  return [
    duration_(r.secondsFromFirstReview), Number(r.secondsFromFirstReview) || 0,
    r.q1Changed ? 'Yes' : 'No',
    (r.q1WordsAdded || 0) + ' added, ' + (r.q1WordsRemoved || 0) + ' removed',
    (bp.text || '') + (bp.voice ? (bp.text ? ' ' : '') + '[voice note]' : ''),
    r.highlightsBefore || 0, r.added || 0, r.removed || 0, r.reworded || 0, r.rehighlighted || 0,
    (r.voiceAdded || 0) + ' added, ' + (r.voiceDeleted || 0) + ' deleted',
    beforeList
  ];
}

// Sheets made by the earlier script lack the review columns: insert them after "Q2: number of highlights".
function upgradeMain_(sheet) {
  const width = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, Math.max(width, 1)).getValues()[0];
  if (header[CORE_HEADERS.length] === REVIEW_HEADERS[0]) return;
  if (header[CORE_HEADERS.length - 1] !== CORE_HEADERS[CORE_HEADERS.length - 1]) return;
  sheet.insertColumnsAfter(CORE_HEADERS.length, REVIEW_HEADERS.length);
  sheet.getRange(1, CORE_HEADERS.length + 1, 1, REVIEW_HEADERS.length)
    .setValues([REVIEW_HEADERS]).setFontWeight('bold').setWrap(true);
}

function upgradeLong_(sheet) {
  const header = sheet.getRange(1, 1, 1, LONG_HEADERS.length).getValues()[0];
  if (header[LONG_HEADERS.length - 1] === LONG_HEADERS[LONG_HEADERS.length - 1]) return;
  if (sheet.getMaxColumns() < LONG_HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), LONG_HEADERS.length - sheet.getMaxColumns());
  sheet.getRange(1, LONG_HEADERS.length, 1, 1).setValues([[LONG_HEADERS[LONG_HEADERS.length - 1]]]).setFontWeight('bold');
}

function blockHeaders_(from, to) {
  const out = [];
  for (let i = from; i <= to; i++) BLOCK.forEach(b => out.push('Q2 highlight ' + i + ': ' + b));
  return out;
}

// Add header columns for highlight blocks the sheet doesn't have yet.
function ensureBlocks_(sheet, count) {
  const have = Math.floor((sheet.getLastColumn() - BASE_HEADERS.length) / BLOCK.length);
  if (count <= have) return;
  const headers = blockHeaders_(have + 1, count);
  const startCol = BASE_HEADERS.length + have * BLOCK.length + 1;
  if (sheet.getMaxColumns() < startCol + headers.length - 1) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), startCol + headers.length - 1 - sheet.getMaxColumns());
  }
  sheet.getRange(1, startCol, 1, headers.length).setValues([headers]).setFontWeight('bold').setWrap(true);
}

function saveVoiceNotes_(data) {
  const links = {};
  (data.voiceNotes || []).forEach(v => {
    const mime = String(v.mime || 'audio/webm').split(';')[0];
    const ext = mime.indexOf('mp4') !== -1 ? 'm4a' : mime.indexOf('ogg') !== -1 ? 'ogg' : 'webm';
    const label = v.id.indexOf('.answer') !== -1 ? 'Q1' : 'Q2';
    const name = [data.code, String(data.name || '').replace(/\s+/g, '-'), label, v.id].join('_') + '.' + ext;
    const blob = Utilities.newBlob(Utilities.base64Decode(v.data), mime, name);
    links[v.id] = getFolder_().createFile(blob).getUrl();
  });
  return links;
}

function getSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setWrap(true);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
  }
  return sh;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('VOICE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* folder was deleted: make a new one */ }
  }
  const folder = DriveApp.createFolder(VOICE_FOLDER_NAME);
  props.setProperty('VOICE_FOLDER_ID', folder.getId());
  return folder;
}

// Stop answers that start with = + - @ from being read as formulas.
function safe_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v == null ? '' : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
