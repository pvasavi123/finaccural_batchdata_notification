const DEFAULT_BATCH_SIZE = 100;
const DEFAULT_WRITE_BATCH_SIZE = 20;

/**
 * Batch size for the *manual* Refresh workflow (see the manual batch
 * queue helpers below) — every Refresh click writes at most this many
 * records to the sheet, regardless of how many new/updated records the
 * backend actually returned in one pull.
 */
export const MANUAL_REFRESH_BATCH_SIZE = 10;

// Background fill applied to rows flagged as newly added since the last
// Master Data Pull. Previously existing rows are left with the default
// (no fill) background.
const NEW_RECORD_FILL_COLOR = "#D9D9D9";

/**
 * Applies (or clears) the "new record" highlight fill across a block of
 * rows, grouping consecutive rows that share the same isNew flag into a
 * single range.format call rather than one per row.
 *
 * @param {Excel.Worksheet} sheet
 * @param {string} startColumn
 * @param {string} endColumn
 * @param {number} batchStartRow - absolute row number of flags[0]
 * @param {boolean[]} flags
 */
function applyNewRecordHighlighting(sheet, startColumn, endColumn, batchStartRow, flags) {
  let i = 0;
  while (i < flags.length) {
    const flagValue = !!flags[i];
    let j = i;
    while (j < flags.length && !!flags[j] === flagValue) j++;

    const rangeStartRow = batchStartRow + i;
    const rangeEndRow = batchStartRow + j - 1;
    const range = sheet.getRange(`${startColumn}${rangeStartRow}:${endColumn}${rangeEndRow}`);

    if (flagValue) {
      range.format.fill.color = NEW_RECORD_FILL_COLOR;
    } else {
      range.format.fill.clear();
    }

    i = j;
  }
}

/**
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {string} startColumn
 * @param {string} endColumn
 * @param {number} startRow
 * @param {any[][]} rows
 * @param {boolean[]|null} [isNewFlags=null] - one boolean per row; true rows get the "new record" fill, false/missing rows get the default (no fill) background
 * @param {number} [batchSize=20]
 */
export async function writeRowsInBatches(
  context,
  sheet,
  startColumn,
  endColumn,
  startRow,
  rows,
  isNewFlags = null,
  batchSize = DEFAULT_WRITE_BATCH_SIZE
) {
  if (!rows || rows.length === 0) return;

  let offset = 0;
  while (offset < rows.length) {
    const batchData = rows.slice(offset, offset + batchSize);
    const batchStartRow = startRow + offset;
    const batchEndRow = batchStartRow + batchData.length - 1;

    try {
      const range = sheet.getRange(`${startColumn}${batchStartRow}:${endColumn}${batchEndRow}`);
      range.values = batchData;

      if (Array.isArray(isNewFlags)) {
        applyNewRecordHighlighting(
          sheet,
          startColumn,
          endColumn,
          batchStartRow,
          isNewFlags.slice(offset, offset + batchData.length)
        );
      }

      await context.sync();
    } catch (error) {
      throw new Error(
        `writeRowsInBatches: failed writing rows ${batchStartRow}-${batchEndRow} to ${startColumn}:${endColumn}: ${error.message || error}`
      );
    }

    offset += batchSize;
  }
}

/**
 * @param {(offset: number, batchSize: number) => Promise<any[]>} fetchBatch
 * @param {Object} [options]
 * @param {number} [options.batchSize=100]
 * @returns {Promise<any[]>}
 */
export async function fetchAllInBatches(fetchBatch, options = {}) {
  const { batchSize = DEFAULT_BATCH_SIZE } = options;

  if (typeof fetchBatch !== "function") {
    throw new Error("fetchAllInBatches: fetchBatch must be a function.");
  }
  if (!Number.isInteger(batchSize) || batchSize <= 0) {
    throw new Error("fetchAllInBatches: batchSize must be a positive integer.");
  }

  const allData = [];
  let offset = 0;

  while (true) {
    let batchData;
    try {
      batchData = await fetchBatch(offset, batchSize);
    } catch (error) {
      throw new Error(`fetchAllInBatches: failed fetching batch at offset ${offset}: ${error.message || error}`);
    }

    if (!Array.isArray(batchData)) {
      throw new Error("fetchAllInBatches: fetchBatch must resolve to an array.");
    }

    if (batchData.length === 0) {
      break;
    }

    allData.push(...batchData);
    offset += batchData.length;
  }

  return allData;
}

/**
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {number} offset
 * @param {number} batchSize
 * @param {Object} [config]
 * @param {number} [config.startRow=1]
 * @param {number} [config.startColumn=0]
 * @param {number} [config.columnCount]
 * @returns {Promise<any[][]>}
 */
export async function readNextBatchFromWorksheet(context, sheet, offset, batchSize, config = {}) {
  const { startRow = 1, startColumn = 0 } = config;
  let { columnCount } = config;

  if (columnCount === undefined) {
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("columnCount");
    await context.sync();
    columnCount = usedRange.isNullObject ? 1 : usedRange.columnCount;
  }

  const absoluteRow = startRow + offset;

  let range;
  try {
    range = sheet.getRangeByIndexes(absoluteRow, startColumn, batchSize, columnCount);
    range.load("values");
    await context.sync();
  } catch (error) {
    throw new Error(
      `readNextBatchFromWorksheet: failed reading rows at offset ${offset}: ${error.message || error}`
    );
  }

  return trimTrailingEmptyRows(range.values);
}

function isRowEmpty(row) {
  return row.every((cell) => cell === "" || cell === null || cell === undefined);
}

function trimTrailingEmptyRows(rows) {
  let end = rows.length;
  while (end > 0 && isRowEmpty(rows[end - 1])) {
    end -= 1;
  }
  return rows.slice(0, end);
}

/**
 * @param {Object} options
 * @param {string} [options.sheetName]
 * @param {number} [options.batchSize=100]
 * @param {number} [options.startRow=1]
 * @param {(allData: any[][]) => (void | Promise<void>)} options.onComplete
 * @returns {Promise<number>}
 */
export async function loadAllWorksheetDataInBatches(options = {}) {
  const { sheetName, batchSize = DEFAULT_BATCH_SIZE, startRow = 1, onComplete } = options;

  if (typeof onComplete !== "function") {
    throw new Error("loadAllWorksheetDataInBatches: options.onComplete must be a function.");
  }

  let allData = [];

  try {
    await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      allData = await fetchAllInBatches(
        (offset, size) => readNextBatchFromWorksheet(context, sheet, offset, size, { startRow }),
        { batchSize }
      );
    });
  } catch (error) {
    if (typeof OfficeExtension !== "undefined" && error instanceof OfficeExtension.Error) {
      console.error("Excel API error:", error.code, error.message, error.debugInfo);
    } else {
      console.error("loadAllWorksheetDataInBatches failed:", error);
    }
    throw error;
  }

  await onComplete(allData);

  return allData.length;
}

/**
 * @param {Excel.RequestContext} context
 * @param {Excel.Worksheet} sheet
 * @param {any[][]} allData
 * @param {Object} [config]
 * @param {number} [config.startRow=1]
 * @param {number} [config.startColumn=0]
 */
export async function writeAllDataToExcelOnce(context, sheet, allData, config = {}) {
  if (!allData || allData.length === 0) return;

  const { startRow = 1, startColumn = 0 } = config;
  const columnCount = allData[0].length;

  try {
    const range = sheet.getRangeByIndexes(startRow, startColumn, allData.length, columnCount);
    range.values = allData;
    await context.sync();
  } catch (error) {
    throw new Error(`writeAllDataToExcelOnce: failed writing ${allData.length} row(s): ${error.message || error}`);
  }
}

/**
 * @param {Object} options
 * @param {(offset: number, batchSize: number) => Promise<any[][]>} options.fetchBatch
 * @param {string} [options.sheetName]
 * @param {number} [options.batchSize=100]
 * @param {number} [options.startRow=1]
 * @param {number} [options.startColumn=0]
 * @returns {Promise<number>}
 */
export async function fetchAllRecordsAndWriteToExcel(options = {}) {
  const { fetchBatch, sheetName, batchSize = DEFAULT_BATCH_SIZE, startRow = 1, startColumn = 0 } = options;

  if (typeof fetchBatch !== "function") {
    throw new Error("fetchAllRecordsAndWriteToExcel: options.fetchBatch must be a function.");
  }

  const allData = await fetchAllInBatches(fetchBatch, { batchSize });

  try {
    await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await writeAllDataToExcelOnce(context, sheet, allData, { startRow, startColumn });
    });
  } catch (error) {
    if (typeof OfficeExtension !== "undefined" && error instanceof OfficeExtension.Error) {
      console.error("Excel API error:", error.code, error.message, error.debugInfo);
    } else {
      console.error("fetchAllRecordsAndWriteToExcel failed:", error);
    }
    throw error;
  }

  return allData.length;
}

// ============================================================
// Manual (click-driven) batch refresh queue
// ============================================================
//
// "Pull Master Data" and "Refresh Schedule" are two triggers for the
// *same* sequential batch loader: whichever one is clicked, it fetches
// (or resumes) the current master-data set and writes the next
// MANUAL_REFRESH_BATCH_SIZE records to the sheet, then advances the
// shared position by that many records — same role as an `FA_NextRowIndex`
// counter in a VBA workbook. The queue below is that shared position: it
// remembers "how far through the current pull did we get" so the *next*
// click — Pull or Refresh, either one — resumes from there instead of
// re-writing (or skipping) records.
//
// The two buttons deliberately share ONE queue per provider/company
// (there is no separate "pull" vs "refresh" position). Clicking either
// button must never rewind that shared position — the position only
// resets when the current queue is fully drained (a brand-new pull
// cycle naturally starts back at record 1) or when the caller explicitly
// clears it (e.g. disconnecting/switching company). Keeping a single
// queue is what guarantees "Pull -> 1-10, Refresh -> 11-20, Pull ->
// 21-30, ..." instead of each button silently restarting the other's
// progress.
//
// Position is kept in localStorage rather than anywhere in the workbook
// itself, matching how the rest of this add-in already persists
// UI/session state across taskpane reloads (see DashboardService's
// STEP_STORAGE_KEY in taskpane.js) — the worksheet's own data is never
// touched to store this bookkeeping, so "keep existing Excel data
// unchanged" holds even for the position marker itself.

const MANUAL_QUEUE_STORAGE_KEY = "fa_manual_batch_queue";

function manualQueueKey(provider, companyId) {
  return `${provider || "unknown"}::${companyId || "unknown"}`;
}

function loadAllManualQueues() {
  try {
    const raw = localStorage.getItem(MANUAL_QUEUE_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch (_) {
    return {};
  }
}

function saveAllManualQueues(all) {
  try {
    localStorage.setItem(MANUAL_QUEUE_STORAGE_KEY, JSON.stringify(all));
  } catch (_) {
    // Storage full/unavailable — the current batch position just won't
    // survive a taskpane reload; this click's own progress still works.
  }
}

/**
 * Reads the shared Pull/Refresh batch position for a provider/company —
 * equivalent to reading `FA_NextRowIndex`. Both buttons call this same
 * getter, so whichever one was clicked last, the other picks up from
 * exactly where it left off.
 * @param {string} provider
 * @param {string} companyId
 * @returns {{records: any[], nextIndex: number, total: number}|null}
 */
export function getManualBatchQueue(provider, companyId) {
  const all = loadAllManualQueues();
  return all[manualQueueKey(provider, companyId)] || null;
}

/**
 * Persists the shared Pull/Refresh batch position — equivalent to
 * writing back an updated `FA_NextRowIndex`. Both buttons call this same
 * setter after appending their batch, so the position they leave behind
 * is visible to the other button's next click too.
 * @param {string} provider
 * @param {string} companyId
 * @param {{records: any[], nextIndex: number, total: number}} queue
 */
export function setManualBatchQueue(provider, companyId, queue) {
  const all = loadAllManualQueues();
  all[manualQueueKey(provider, companyId)] = queue;
  saveAllManualQueues(all);
}

/**
 * Drops the shared stored batch position for a provider/company — call
 * this once the queue is fully drained (the natural end of a pull
 * cycle, so the next click starts a fresh cycle back at record 1), or
 * when the underlying data it was sliced from is no longer valid (e.g.
 * disconnecting or switching company). Neither Pull Master Data nor
 * Refresh should call this just because *it* was the button clicked —
 * only when the shared queue itself is actually finished or invalidated.
 * @param {string} provider
 * @param {string} companyId
 */
export function clearManualBatchQueue(provider, companyId) {
  const all = loadAllManualQueues();
  delete all[manualQueueKey(provider, companyId)];
  saveAllManualQueues(all);
}

/**
 * Slices the next up-to-`batchSize` not-yet-written records off a manual
 * refresh queue. Pure function — does not touch storage; callers persist
 * (or clear) the queue themselves based on the returned `isDone` flag.
 *
 * @param {any[]} records - full ordered list discovered by the last pull
 * @param {number} nextIndex - how many records have already been written
 * @param {number} [batchSize=MANUAL_REFRESH_BATCH_SIZE]
 * @returns {{batch: any[], nextIndex: number, isDone: boolean, total: number, batchStart: number, batchEnd: number}}
 */
export function takeNextManualBatch(records, nextIndex, batchSize = MANUAL_REFRESH_BATCH_SIZE) {
  const list = Array.isArray(records) ? records : [];
  const start = Math.max(0, Number.isInteger(nextIndex) ? nextIndex : 0);
  const batch = list.slice(start, start + batchSize);

  return {
    batch,
    nextIndex: start + batch.length,
    isDone: start + batch.length >= list.length,
    total: list.length,
    // 1-indexed, for "rows 101-200 of 437" style progress messaging.
    batchStart: list.length === 0 ? 0 : start + 1,
    batchEnd: start + batch.length
  };
}
