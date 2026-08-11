'use strict';

/**
 * Excel Schema Validation
 * -----------------------------------------------------------------
 * Validates workbook STRUCTURE only — which sheets exist, which
 * required columns each sheet has, and which columns are unexpected.
 * Does not look at cell values (that's dataTypeValidator.js).
 * -----------------------------------------------------------------
 */

function normalizeHeader(header) {
    return String(header).trim().toLowerCase();
}

/**
 * Validate one already-parsed worksheet's headers against a schema.
 * @param {string} sheetName
 * @param {string[]} actualHeaders
 * @param {{columns: Array<{key:string, header:string, required:boolean}>}} schema
 * @returns {{ sheet: string, ok: boolean, missingColumns: string[], unexpectedColumns: string[] }}
 */
function validateSheetSchema(sheetName, actualHeaders, schema) {
    const actualSet = new Set((actualHeaders || []).map(normalizeHeader));
    const expectedSet = new Set(schema.columns.map((c) => normalizeHeader(c.header)));

    const missingColumns = schema.columns
        .filter((c) => c.required && !actualSet.has(normalizeHeader(c.header)))
        .map((c) => c.header);

    const unexpectedColumns = (actualHeaders || []).filter((h) => !expectedSet.has(normalizeHeader(h)));

    return {
        sheet: sheetName,
        ok: missingColumns.length === 0,
        missingColumns,
        unexpectedColumns
    };
}

/**
 * Validate an entire parsed workbook against a set of expected sheets.
 * @param {{ sheetNames: string[], sheets: Object }} parsed - output of ExcelParser.parseWorkbook
 * @param {Object<string, {sheetName:string, columns:Array}>} sheetSchemas - name -> schema
 * @param {string[]} [optionalSheets] - sheet names allowed to be absent entirely
 * @returns {{ ok: boolean, missingSheets: string[], unexpectedSheets: string[], sheetResults: Array }}
 */
function validateWorkbookSchema(parsed, sheetSchemas, optionalSheets = []) {
    const expectedSheetNames = Object.keys(sheetSchemas);
    const actualSheetSet = new Set(parsed.sheetNames);
    const expectedSheetSet = new Set(expectedSheetNames);

    const missingSheets = expectedSheetNames.filter(
        (name) => !actualSheetSet.has(name) && !optionalSheets.includes(name)
    );
    const unexpectedSheets = parsed.sheetNames.filter((name) => !expectedSheetSet.has(name));

    const sheetResults = expectedSheetNames
        .filter((name) => actualSheetSet.has(name))
        .map((name) => validateSheetSchema(name, parsed.sheets[name].headers, sheetSchemas[name]));

    const ok = missingSheets.length === 0 && sheetResults.every((r) => r.ok);

    return { ok, missingSheets, unexpectedSheets, sheetResults };
}

module.exports = { validateSheetSchema, validateWorkbookSchema, normalizeHeader };
