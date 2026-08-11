const { validateRow, validateSheetDataTypes } = require('../../../../src/modules/excelValidation/validators/dataTypeValidator');
const { SHEET_SCHEMAS } = require('../../../../src/modules/excelValidation/schemas/masterDataSchemas');

describe('dataTypeValidator', () => {
    describe('validateRow', () => {
        it('returns no errors for a fully valid Customers row', () => {
            const errors = validateRow(SHEET_SCHEMAS.Customers.columns, {
                ID: 'C1', Name: 'Acme Co', 'Company Name': 'Acme', Email: 'ap@acme.com', Balance: 150.5
            });
            expect(errors).toEqual([]);
        });

        it('flags a required field that is empty', () => {
            const errors = validateRow(SHEET_SCHEMAS.Customers.columns, {
                ID: '', Name: 'Acme Co', 'Company Name': 'Acme', Email: '', Balance: 150.5
            });
            expect(errors).toEqual([expect.objectContaining({ field: 'id', reason: 'Required field is empty' })]);
        });

        it('flags a non-numeric Balance', () => {
            const errors = validateRow(SHEET_SCHEMAS.Customers.columns, {
                ID: 'C1', Name: 'Acme Co', 'Company Name': 'Acme', Email: 'ap@acme.com', Balance: 'not-a-number'
            });
            expect(errors).toEqual([expect.objectContaining({ field: 'balance' })]);
        });

        it('flags a malformed email', () => {
            const errors = validateRow(SHEET_SCHEMAS.Customers.columns, {
                ID: 'C1', Name: 'Acme Co', 'Company Name': 'Acme', Email: 'not-an-email', Balance: 10
            });
            expect(errors).toEqual([expect.objectContaining({ field: 'email' })]);
        });

        it('allows an optional email to be blank', () => {
            const errors = validateRow(SHEET_SCHEMAS.Customers.columns, {
                ID: 'C1', Name: 'Acme Co', 'Company Name': 'Acme', Email: '', Balance: 10
            });
            expect(errors).toEqual([]);
        });

        it('flags a Status value outside the enum', () => {
            const errors = validateRow(SHEET_SCHEMAS.Classes.columns, {
                ID: 'CL1', Name: 'Marketing', Status: 'Archived'
            });
            expect(errors).toEqual([expect.objectContaining({ field: 'status' })]);
        });

        it('accepts enum values case-insensitively', () => {
            const errors = validateRow(SHEET_SCHEMAS.Classes.columns, {
                ID: 'CL1', Name: 'Marketing', Status: 'active'
            });
            expect(errors).toEqual([]);
        });
    });

    describe('validateSheetDataTypes', () => {
        it('aggregates per-row results for a sheet', () => {
            const parsedSheet = {
                headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'],
                rows: [
                    { __rowNumber: 2, values: { ID: 'C1', Name: 'Good', 'Company Name': '', Email: '', Balance: 10 } },
                    { __rowNumber: 3, values: { ID: '', Name: 'Bad', 'Company Name': '', Email: 'x', Balance: 'oops' } }
                ]
            };

            const result = validateSheetDataTypes('Customers', parsedSheet, SHEET_SCHEMAS.Customers);
            expect(result.totalRows).toBe(2);
            expect(result.validRows).toBe(1);
            expect(result.invalidRows).toBe(1);
            expect(result.rowResults[1].ok).toBe(false);
            expect(result.rowResults[1].errors.length).toBeGreaterThanOrEqual(2); // id required + email + balance
        });
    });
});
