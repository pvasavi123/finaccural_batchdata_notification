'use strict';

const ExcelValidationService = require('./service');
const { ValidationError } = require('../../core/errors/AppError');

const MAX_FILE_BYTES = 10 * 1024 * 1024; // 10MB decoded — generous for a master-data workbook

/**
 * ExcelValidationController
 * -----------------------------------------------------------------
 * The workbook travels as base64 inside the JSON request body
 * (`fileBase64`) rather than multipart/form-data — this API's global
 * `validateContentType` middleware (core/middleware/validateHeaders.js)
 * rejects any POST/PUT/PATCH that isn't `application/json`, and adding
 * a multipart exception there would widen that Header Validation gate
 * for every route, not just this module's. Base64-in-JSON keeps this
 * module's file upload on the same contract as every other endpoint in
 * this API.
 * -----------------------------------------------------------------
 */
class ExcelValidationController {

    /** Decode req.body.fileBase64 -> Buffer, with a friendly error and a size cap. */
    decodeFile(req) {
        const { fileBase64 } = req.body;
        if (!fileBase64 || typeof fileBase64 !== 'string') {
            throw new ValidationError('fileBase64 is required — base64-encode the .xlsx file and send it as a JSON string.');
        }

        const cleaned = fileBase64.includes(',') ? fileBase64.slice(fileBase64.indexOf(',') + 1) : fileBase64;
        const buffer = Buffer.from(cleaned, 'base64');

        if (!buffer.length) {
            throw new ValidationError('fileBase64 did not decode to any bytes — check the encoding.');
        }
        if (buffer.length > MAX_FILE_BYTES) {
            throw new ValidationError(`Uploaded file is too large (${buffer.length} bytes). Max allowed is ${MAX_FILE_BYTES} bytes.`);
        }

        return buffer;
    }

    /**
     * POST /api/excel-validation/schema-check
     * Body: { fileBase64, schema }
     */
    validateSchema = async (req, res, next) => {
        try {
            const buffer = this.decodeFile(req);
            const { schemaResult } = await ExcelValidationService.validateSchema(buffer, req.body.schema);
            res.json({ schema: req.body.schema, ...schemaResult });
        } catch (err) {
            next(err);
        }
    };

    /**
     * POST /api/excel-validation/data-type-check
     * Body: { fileBase64, schema }
     */
    validateDataTypes = async (req, res, next) => {
        try {
            const buffer = this.decodeFile(req);
            const { dataTypeResult } = await ExcelValidationService.validateDataTypes(buffer, req.body.schema);
            res.json({ schema: req.body.schema, ...dataTypeResult });
        } catch (err) {
            next(err);
        }
    };

    /**
     * POST /api/excel-validation/report?format=xlsx
     * Body: { fileBase64, schema }
     * Default response is the PASS/FAIL JSON report; ?format=xlsx
     * streams back a Summary+Failures .xlsx workbook instead.
     */
    generateReport = async (req, res, next) => {
        try {
            const buffer = this.decodeFile(req);
            const wantsWorkbook = String(req.query.format || '').toLowerCase() === 'xlsx';

            if (wantsWorkbook) {
                const { report, workbookBuffer } = await ExcelValidationService.generateReportWorkbook(buffer, req.body.schema);
                res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                res.setHeader('Content-Disposition', 'attachment; filename="excel_validation_report.xlsx"');
                res.setHeader('X-Report-Passed', String(report.passed));
                res.setHeader('X-Report-Failed', String(report.failed));
                return res.send(workbookBuffer);
            }

            const { report } = await ExcelValidationService.generateReport(buffer, req.body.schema);
            return res.json({ schema: req.body.schema, ...report });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * POST /api/excel-validation/vs-api
     * Body: { fileBase64, sheet, platform }
     */
    compareWithApi = async (req, res, next) => {
        try {
            const buffer = this.decodeFile(req);
            const { sheet, platform } = req.body;
            const mail = req.user.email; // never trust a client-suppliable mail — same rule as every other module
            const { diff, keyField } = await ExcelValidationService.compareWithApi(buffer, sheet, platform, mail);
            res.json({ sheet, platform, keyField, ...diff });
        } catch (err) {
            next(err);
        }
    };

    /**
     * POST /api/excel-validation/vs-database
     * Body: { fileBase64 } — expects a "Connections" sheet
     */
    compareWithDatabase = async (req, res, next) => {
        try {
            const buffer = this.decodeFile(req);
            const mail = req.user.email;
            const { diff, keyField } = await ExcelValidationService.compareWithDatabase(buffer, mail);
            res.json({ sheet: 'Connections', keyField, ...diff });
        } catch (err) {
            next(err);
        }
    };
}

module.exports = new ExcelValidationController();
