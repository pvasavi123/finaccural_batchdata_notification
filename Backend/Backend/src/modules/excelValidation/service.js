'use strict';

const ExcelParser = require('./parser');
const { validateWorkbookSchema } = require('./validators/schemaValidator');
const { validateWorkbookDataTypes } = require('./validators/dataTypeValidator');
const { buildReport, reportToWorkbookBuffer } = require('./report/reportGenerator');
const { compareSheetWithApi } = require('./compare/apiComparator');
const { compareConnectionsWithDatabase } = require('./compare/dbComparator');
const {
    SHEET_SCHEMAS,
    MASTER_DATA_WORKBOOK_SHEETS,
    OPTIONAL_WORKBOOK_SHEETS,
    API_COMPARABLE_SHEETS,
    getSchema
} = require('./schemas/masterDataSchemas');
const { ValidationError } = require('../../core/errors/AppError');

const MASTER_DATA_KEY = 'MasterData';

/**
 * ExcelValidationService
 * -----------------------------------------------------------------
 * Orchestrates the five capabilities this module provides:
 *   1. Excel Schema Validation      -> validateSchema()
 *   2. Excel Data-Type Validation   -> validateDataTypes()
 *   3. Excel PASS/FAIL Report       -> generateReport() / generateReportWorkbook()
 *   4. Excel vs API Validation      -> compareWithApi()
 *   5. Excel vs Database Reconciliation -> compareWithDatabase()
 *
 * Every method takes the uploaded workbook as a Buffer (the controller
 * is responsible for base64-decoding the request body into one) and a
 * `schemaKey` identifying which contract to validate against — either
 * one of the individual sheet names in masterDataSchemas.js
 * (Company/Customers/Vendors/Accounts/Classes/Locations/Connections),
 * or the special "MasterData" key which validates the full multi-sheet
 * workbook produced by /api/quickbooks/export or /api/xero/export at
 * once.
 * -----------------------------------------------------------------
 */
class ExcelValidationService {

    /**
     * @param {string} schemaKey
     * @returns {{ schemas: Object, optionalSheets: string[] }}
     */
    static resolveSchemas(schemaKey) {
        if (schemaKey === MASTER_DATA_KEY) {
            const schemas = {};
            MASTER_DATA_WORKBOOK_SHEETS.forEach((name) => { schemas[name] = SHEET_SCHEMAS[name]; });
            return { schemas, optionalSheets: OPTIONAL_WORKBOOK_SHEETS };
        }

        return { schemas: { [schemaKey]: getSchema(schemaKey) }, optionalSheets: [] };
    }

    /**
     * Excel Schema Validation.
     * @param {Buffer} buffer
     * @param {string} schemaKey
     */
    static async validateSchema(buffer, schemaKey) {
        const parsed = await ExcelParser.parseWorkbook(buffer);
        const { schemas, optionalSheets } = ExcelValidationService.resolveSchemas(schemaKey);
        return { parsed, schemaResult: validateWorkbookSchema(parsed, schemas, optionalSheets) };
    }

    /**
     * Excel Data-Type Validation.
     * @param {Buffer} buffer
     * @param {string} schemaKey
     */
    static async validateDataTypes(buffer, schemaKey) {
        const parsed = await ExcelParser.parseWorkbook(buffer);
        const { schemas } = ExcelValidationService.resolveSchemas(schemaKey);
        return { parsed, dataTypeResult: validateWorkbookDataTypes(parsed, schemas) };
    }

    /**
     * Excel PASS/FAIL Report — runs schema + data-type validation
     * together and aggregates into totals/pass-rate/failure reasons.
     * @param {Buffer} buffer
     * @param {string} schemaKey
     */
    static async generateReport(buffer, schemaKey) {
        const parsed = await ExcelParser.parseWorkbook(buffer);
        const { schemas, optionalSheets } = ExcelValidationService.resolveSchemas(schemaKey);

        const schemaResult = validateWorkbookSchema(parsed, schemas, optionalSheets);
        const dataTypeResult = validateWorkbookDataTypes(parsed, schemas);
        const report = buildReport({ schemaResult, dataTypeResult });

        return { parsed, schemaResult, dataTypeResult, report };
    }

    /**
     * Same as generateReport(), but also renders the report as an .xlsx
     * workbook buffer (Summary + Failures sheets) ready to stream back
     * as a download.
     * @param {Buffer} buffer
     * @param {string} schemaKey
     */
    static async generateReportWorkbook(buffer, schemaKey) {
        const { report } = await ExcelValidationService.generateReport(buffer, schemaKey);
        const workbookBuffer = await reportToWorkbookBuffer(report);
        return { report, workbookBuffer };
    }

    /**
     * Excel vs API Validation — diffs one sheet's Excel rows against the
     * same data pulled live from QuickBooks/Xero right now.
     * @param {Buffer} buffer
     * @param {string} sheetName
     * @param {'quickbooks'|'xero'} platform
     * @param {string} mail - authenticated user's email
     */
    static async compareWithApi(buffer, sheetName, platform, mail) {
        if (!API_COMPARABLE_SHEETS.includes(sheetName)) {
            throw new ValidationError(
                `"${sheetName}" cannot be compared against a live API.`,
                `Expected one of: ${API_COMPARABLE_SHEETS.join(', ')}`
            );
        }

        const parsed = await ExcelParser.parseWorkbook(buffer);
        const parsedSheet = parsed.sheets[sheetName];
        if (!parsedSheet) {
            throw new ValidationError(`The uploaded workbook has no "${sheetName}" sheet.`);
        }

        const result = await compareSheetWithApi({ sheetName, parsedSheet, platform, mail });
        return { parsed, ...result };
    }

    /**
     * Excel vs Database Reconciliation — diffs the "Connections" sheet
     * against this user's live QuickBooksToken/XeroToken rows.
     * @param {Buffer} buffer
     * @param {string} mail - authenticated user's email
     */
    static async compareWithDatabase(buffer, mail) {
        const parsed = await ExcelParser.parseWorkbook(buffer);
        const parsedSheet = parsed.sheets.Connections;
        if (!parsedSheet) {
            throw new ValidationError('The uploaded workbook has no "Connections" sheet.');
        }

        const result = await compareConnectionsWithDatabase({ parsedSheet, mail });
        return { parsed, ...result };
    }
}

module.exports = ExcelValidationService;
module.exports.MASTER_DATA_KEY = MASTER_DATA_KEY;
