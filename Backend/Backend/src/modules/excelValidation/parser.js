'use strict';

const ExcelJS = require('exceljs');
const { ValidationError } = require('../../core/errors/AppError');
const { readCellValue, isEmpty } = require('./validators/typeCheckers');

/**
 * ExcelParser
 * -----------------------------------------------------------------
 * Turns an uploaded .xlsx buffer into a plain-JS structure every other
 * piece of this module (schema validator, data-type validator,
 * comparators) works against, so none of them need to touch exceljs
 * directly.
 * -----------------------------------------------------------------
 */
class ExcelParser {

    /**
     * @param {Buffer} buffer
     * @returns {Promise<{ sheetNames: string[], sheets: Object<string, { headers: string[], rows: Array<{ __rowNumber: number, values: Object }> }> }>}
     */
    static async parseWorkbook(buffer) {
        if (!buffer || !buffer.length) {
            throw new ValidationError('Uploaded file is empty.');
        }

        const workbook = new ExcelJS.Workbook();
        try {
            await workbook.xlsx.load(buffer);
        } catch (err) {
            throw new ValidationError('Could not read the uploaded file — is it a valid .xlsx workbook?', err.message);
        }

        const sheets = {};
        const sheetNames = [];

        workbook.eachSheet((worksheet) => {
            sheetNames.push(worksheet.name);
            sheets[worksheet.name] = ExcelParser.parseWorksheet(worksheet);
        });

        return { sheetNames, sheets };
    }

    /**
     * @param {import('exceljs').Worksheet} worksheet
     * @returns {{ headers: string[], rows: Array<{ __rowNumber: number, values: Object }> }}
     */
    static parseWorksheet(worksheet) {
        const headerRow = worksheet.getRow(1);
        const headers = [];

        headerRow.eachCell({ includeEmpty: false }, (cell) => {
            const value = readCellValue(cell.value);
            if (!isEmpty(value)) {
                headers.push(String(value).trim());
            }
        });

        const rows = [];
        const lastRow = worksheet.rowCount;

        for (let rowNumber = 2; rowNumber <= lastRow; rowNumber++) {
            const row = worksheet.getRow(rowNumber);

            // Skip fully empty rows (exceljs can leave gaps after
            // deletions, or trailing rows with only formatting).
            let hasValue = false;
            const values = {};

            headers.forEach((header, index) => {
                const cell = row.getCell(index + 1);
                const value = readCellValue(cell.value);
                values[header] = value;
                if (!isEmpty(value)) hasValue = true;
            });

            if (hasValue) {
                rows.push({ __rowNumber: rowNumber, values });
            }
        }

        return { headers, rows };
    }
}

module.exports = ExcelParser;
