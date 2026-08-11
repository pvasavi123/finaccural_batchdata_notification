const { validateSheetSchema, validateWorkbookSchema } = require('../../../../src/modules/excelValidation/validators/schemaValidator');
const { SHEET_SCHEMAS, MASTER_DATA_WORKBOOK_SHEETS, OPTIONAL_WORKBOOK_SHEETS } = require('../../../../src/modules/excelValidation/schemas/masterDataSchemas');

describe('schemaValidator', () => {
    describe('validateSheetSchema', () => {
        it('passes when every required column is present (extra/case-insensitive headers ignored)', () => {
            const result = validateSheetSchema('Customers', ['id', 'NAME', 'Company Name', 'Email', 'Balance'], SHEET_SCHEMAS.Customers);
            expect(result.ok).toBe(true);
            expect(result.missingColumns).toEqual([]);
        });

        it('flags missing required columns', () => {
            const result = validateSheetSchema('Customers', ['ID', 'Company Name'], SHEET_SCHEMAS.Customers);
            expect(result.ok).toBe(false);
            expect(result.missingColumns).toEqual(expect.arrayContaining(['Name', 'Balance']));
        });

        it('flags unexpected columns without failing the sheet', () => {
            const result = validateSheetSchema('Customers', ['ID', 'Name', 'Balance', 'Some Random Column'], SHEET_SCHEMAS.Customers);
            expect(result.ok).toBe(true);
            expect(result.unexpectedColumns).toEqual(['Some Random Column']);
        });
    });

    describe('validateWorkbookSchema', () => {
        const buildParsed = (sheets) => ({ sheetNames: Object.keys(sheets), sheets });

        it('passes a full valid master-data workbook', () => {
            const parsed = buildParsed({
                Company:   { headers: ['ID', 'Company Name', 'Legal Name'] },
                Customers: { headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'] },
                Vendors:   { headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'] },
                Accounts:  { headers: ['ID', 'Acct #', 'Name', 'Account Type', 'Sub Type', 'Balance'] },
                Classes:   { headers: ['ID', 'Name', 'Status'] },
                Locations: { headers: ['ID', 'Name', 'Status'] }
            });

            const schemas = {};
            MASTER_DATA_WORKBOOK_SHEETS.forEach((n) => { schemas[n] = SHEET_SCHEMAS[n]; });

            const result = validateWorkbookSchema(parsed, schemas, OPTIONAL_WORKBOOK_SHEETS);
            expect(result.ok).toBe(true);
            expect(result.missingSheets).toEqual([]);
        });

        it('treats Company as optional when absent', () => {
            const parsed = buildParsed({
                Customers: { headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'] },
                Vendors:   { headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'] },
                Accounts:  { headers: ['ID', 'Acct #', 'Name', 'Account Type', 'Sub Type', 'Balance'] },
                Classes:   { headers: ['ID', 'Name', 'Status'] },
                Locations: { headers: ['ID', 'Name', 'Status'] }
            });

            const schemas = {};
            MASTER_DATA_WORKBOOK_SHEETS.forEach((n) => { schemas[n] = SHEET_SCHEMAS[n]; });

            const result = validateWorkbookSchema(parsed, schemas, OPTIONAL_WORKBOOK_SHEETS);
            expect(result.ok).toBe(true);
            expect(result.missingSheets).toEqual([]);
        });

        it('reports a missing required sheet', () => {
            const parsed = buildParsed({
                Customers: { headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'] }
            });

            const schemas = {};
            MASTER_DATA_WORKBOOK_SHEETS.forEach((n) => { schemas[n] = SHEET_SCHEMAS[n]; });

            const result = validateWorkbookSchema(parsed, schemas, OPTIONAL_WORKBOOK_SHEETS);
            expect(result.ok).toBe(false);
            expect(result.missingSheets).toEqual(expect.arrayContaining(['Vendors', 'Accounts', 'Classes', 'Locations']));
        });
    });
});
