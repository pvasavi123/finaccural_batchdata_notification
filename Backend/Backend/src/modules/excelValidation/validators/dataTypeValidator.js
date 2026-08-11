'use strict';

const { checkColumnValue } = require('./typeCheckers');

/**
 * Excel Data-Type Validation
 * -----------------------------------------------------------------
 * Validates cell VALUES (type + format) for every row of an
 * already-parsed worksheet, against the same column schema used by
 * schemaValidator.js. Runs independently of schema validation so a
 * caller can run either check alone, or both (see reportGenerator.js).
 *
 * Only columns declared in the schema are checked — an unexpected
 * column (already flagged by schemaValidator) is left alone here.
 * A row missing a declared column entirely is treated as an empty
 * value for that column (required-ness is enforced the same way).
 * -----------------------------------------------------------------
 */

/**
 * @param {{ header: string, rowNumber: number }} ctx
 * @param {Array<{key:string, header:string, type:string, required:boolean, enumValues?:string[]}>} columns
 * @param {Object} rowValues - header -> raw cell value, as produced by ExcelParser
 * @returns {Array<{ field: string, header: string, reason: string, value: * }>}
 */
function validateRow(columns, rowValues) {
    const errors = [];

    columns.forEach((column) => {
        const rawValue = rowValues[column.header];
        const result = checkColumnValue(rawValue, column);
        if (!result.valid) {
            errors.push({
                field:  column.key,
                header: column.header,
                reason: result.reason,
                value:  rawValue === undefined ? null : rawValue
            });
        }
    });

    return errors;
}

/**
 * @param {string} sheetName
 * @param {{ headers: string[], rows: Array<{__rowNumber:number, values:Object}> }} parsedSheet
 * @param {{ columns: Array }} schema
 * @returns {{ sheet: string, totalRows: number, validRows: number, invalidRows: number, rowResults: Array }}
 */
function validateSheetDataTypes(sheetName, parsedSheet, schema) {
    const rowResults = parsedSheet.rows.map((row) => {
        const errors = validateRow(schema.columns, row.values);
        return { rowNumber: row.__rowNumber, ok: errors.length === 0, errors };
    });

    return {
        sheet: sheetName,
        totalRows: rowResults.length,
        validRows: rowResults.filter((r) => r.ok).length,
        invalidRows: rowResults.filter((r) => !r.ok).length,
        rowResults
    };
}

/**
 * @param {{ sheetNames: string[], sheets: Object }} parsed
 * @param {Object<string, {columns:Array}>} sheetSchemas
 * @returns {{ sheetResults: Array }}
 */
function validateWorkbookDataTypes(parsed, sheetSchemas) {
    const sheetResults = Object.keys(sheetSchemas)
        .filter((name) => parsed.sheets[name])
        .map((name) => validateSheetDataTypes(name, parsed.sheets[name], sheetSchemas[name]));

    return { sheetResults };
}

module.exports = { validateRow, validateSheetDataTypes, validateWorkbookDataTypes };
