/**
 * Weekly highlights auto-updater.
 *
 * Adds a new weekly block (merged "Week" cell + N empty rows) to the
 * `highlights` tab. Designed to be installed as a time-based trigger
 * that runs once a week.
 *
 * Setup:
 *   1. Open the spreadsheet → Extensions → Apps Script.
 *   2. Paste this file's contents into `Code.gs`.
 *   3. Run `installWeeklyTrigger` once and grant permissions.
 *      That schedules `addWeeklyRows` to run every Sunday morning.
 *   4. (Optional) Run `addWeeklyRows` manually to test.
 */

const CONFIG = {
  sheetName: 'highlights',
  weekColumn: 1,        // Column A
  totalColumns: 6,      // A..F (Week, Operation, Brief, Source Link, Views, Engagment)
  rowsPerWeek: 6,       // How many blank rows to allocate per week
  triggerWeekday: ScriptApp.WeekDay.SUNDAY,
  triggerHour: 6,       // Local time of the Apps Script project
  // Israeli work-week: Sunday → Saturday. Override if you want a different cadence.
  weekStartsOn: 0,      // 0 = Sunday, 1 = Monday, ...
};

/**
 * Entry point for the time-based trigger.
 */
function addWeeklyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet bound to this script.');
  }
  Logger.log('Spreadsheet: %s (%s)', ss.getName(), ss.getUrl());

  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    const available = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(
      'Sheet "' + CONFIG.sheetName + '" not found. Available tabs: ' + available
    );
  }

  const layout = scanWeekBlocks(sheet);
  const nextRange = computeNextWeekRange(layout);
  const label = formatWeekLabel(nextRange.start, nextRange.end);

  const existing = layout.blocks.find(b => b.label === label);
  if (existing) {
    Logger.log('Week %s already exists at row %s, skipping insert.', label, existing.startRow);
    focusBlock(sheet, existing.startRow, existing.endRow - existing.startRow + 1);
    return;
  }

  const insertedAt = insertWeekBlock(sheet, layout, label);
  Logger.log('Inserted week block "%s" at row %s on tab "%s".', label, insertedAt, sheet.getName());
  focusBlock(sheet, insertedAt, CONFIG.rowsPerWeek);
}

/**
 * Jump the active view to the new block so it's obvious where it landed.
 */
function focusBlock(sheet, startRow, numRows) {
  try {
    sheet.activate();
    const range = sheet.getRange(startRow, 1, numRows, CONFIG.totalColumns);
    sheet.setActiveRange(range);
  } catch (e) {
    Logger.log('focusBlock skipped: %s', e);
  }
}

/**
 * Diagnostic: log which spreadsheet + tabs this script is bound to.
 * Run this once if `addWeeklyRows` reports success but you don't see changes.
 */
function whereAmI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('No active spreadsheet. This script is probably standalone, not bound to a sheet.');
    return;
  }
  Logger.log('Name: %s', ss.getName());
  Logger.log('URL:  %s', ss.getUrl());
  Logger.log('Tabs: %s', ss.getSheets().map(s => s.getName()).join(' | '));
  Logger.log('Target tab "%s" exists: %s', CONFIG.sheetName, !!ss.getSheetByName(CONFIG.sheetName));
}

/**
 * Install / refresh the weekly trigger. Run this once after pasting the script.
 */
function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'addWeeklyRows') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('addWeeklyRows')
    .timeBased()
    .onWeekDay(CONFIG.triggerWeekday)
    .atHour(CONFIG.triggerHour)
    .create();
  Logger.log('Weekly trigger installed.');
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

/**
 * Scan column A and return { headerRow, blocks: [{label, startRow, endRow, start, end}] }
 * in the order they appear in the sheet.
 */
function scanWeekBlocks(sheet) {
  const lastRow = sheet.getLastRow();
  const headerRow = findHeaderRow(sheet, lastRow);
  const blocks = [];

  if (lastRow <= headerRow) {
    return { headerRow: headerRow, blocks: blocks, lastRow: lastRow };
  }

  const values = sheet
    .getRange(headerRow + 1, CONFIG.weekColumn, lastRow - headerRow, 1)
    .getValues();

  let current = null;
  for (let i = 0; i < values.length; i++) {
    const row = headerRow + 1 + i;
    const raw = values[i][0];
    const label = raw === null || raw === undefined ? '' : String(raw).trim();
    if (label) {
      if (current) blocks.push(current);
      const parsed = parseWeekLabel(label);
      current = {
        label: label,
        startRow: row,
        endRow: row,
        start: parsed ? parsed.start : null,
        end: parsed ? parsed.end : null,
      };
    } else if (current) {
      current.endRow = row;
    }
  }
  if (current) blocks.push(current);

  return { headerRow: headerRow, blocks: blocks, lastRow: lastRow };
}

/**
 * The header row is the row that contains "Week" in column A.
 * Falls back to row 2 (matches the screenshot) if not found.
 */
function findHeaderRow(sheet, lastRow) {
  const scan = Math.min(lastRow, 10);
  if (scan === 0) return 2;
  const values = sheet.getRange(1, CONFIG.weekColumn, scan, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === 'week') return i + 1;
  }
  return 2;
}

/**
 * Return the Sun-Sat range of the week that contains "today".
 *
 * We deliberately ignore whatever is already in the sheet and anchor on
 * today's date: the goal is "add the current week each Sunday", and that's
 * robust against weird labels or stale historical data.
 */
function computeNextWeekRange(_layout) {
  const start = startOfWeek(new Date(), CONFIG.weekStartsOn);
  const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
  return { start: start, end: end };
}

function startOfWeek(date, weekStartsOn) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diff = (d.getDay() - weekStartsOn + 7) % 7;
  d.setDate(d.getDate() - diff);
  return d;
}

/**
 * Insert a new week block at the top of the data area, directly under
 * the header row. Existing blocks shift down.
 */
function insertWeekBlock(sheet, layout, label) {
  const rows = CONFIG.rowsPerWeek;
  const insertAt = layout.headerRow + 1;
  sheet.insertRowsBefore(insertAt, rows);

  // Defensive: existing merges in adjacent rows can absorb the new rows
  // when we insertRowsBefore. Break any merge that touches our target
  // range before re-merging it cleanly.
  const blockRange = sheet.getRange(insertAt, 1, rows, CONFIG.totalColumns);
  blockRange.breakApart();

  const weekRange = sheet.getRange(insertAt, CONFIG.weekColumn, rows, 1);
  weekRange.merge();
  weekRange
    .setValue(label)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold');

  sheet
    .getRange(insertAt, CONFIG.weekColumn, rows, CONFIG.totalColumns)
    .setBorder(true, true, true, true, true, true);

  return insertAt;
}

/**
 * Parse week labels like "3-9.5" or "26.4-2.5".
 * Returns { start: Date, end: Date } or null.
 */
function parseWeekLabel(label) {
  if (!label) return null;
  const year = new Date().getFullYear();

  // Same-month form: "3-9.5"
  let m = label.match(/^\s*(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\s*$/);
  if (m) {
    const month = +m[3];
    return {
      start: new Date(year, month - 1, +m[1]),
      end: new Date(year, month - 1, +m[2]),
    };
  }

  // Cross-month form: "26.4-2.5"
  m = label.match(/^\s*(\d{1,2})\.(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\s*$/);
  if (m) {
    let startMonth = +m[2];
    let endMonth = +m[4];
    let startYear = year;
    let endYear = year;
    // Handle year wrap (Dec → Jan)
    if (endMonth < startMonth) endYear = year + 1;
    return {
      start: new Date(startYear, startMonth - 1, +m[1]),
      end: new Date(endYear, endMonth - 1, +m[3]),
    };
  }

  return null;
}

function formatWeekLabel(start, end) {
  const sd = start.getDate();
  const sm = start.getMonth() + 1;
  const ed = end.getDate();
  const em = end.getMonth() + 1;
  if (sm === em) return sd + '-' + ed + '.' + em;
  return sd + '.' + sm + '-' + ed + '.' + em;
}
