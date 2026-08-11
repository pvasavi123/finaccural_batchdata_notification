const { buildReport, reportToWorkbookBuffer } = require('../../../../src/modules/excelValidation/report/reportGenerator');

describe('reportGenerator', () => {
    describe('buildReport', () => {
        it('reports 100% pass rate when there are no schema or data-type problems', () => {
            const schemaResult = {
                ok: true,
                missingSheets: [],
                unexpectedSheets: [],
                sheetResults: [{ sheet: 'Customers', ok: true, missingColumns: [], unexpectedColumns: [] }]
            };
            const dataTypeResult = {
                sheetResults: [{
                    sheet: 'Customers',
                    totalRows: 3,
                    validRows: 3,
                    invalidRows: 0,
                    rowResults: [
                        { rowNumber: 2, ok: true, errors: [] },
                        { rowNumber: 3, ok: true, errors: [] },
                        { rowNumber: 4, ok: true, errors: [] }
                    ]
                }]
            };

            const report = buildReport({ schemaResult, dataTypeResult });
            expect(report.totalRecords).toBe(3);
            expect(report.passed).toBe(3);
            expect(report.failed).toBe(0);
            expect(report.passRate).toBe(1);
            expect(report.failures).toEqual([]);
        });

        it('counts every row of a structurally-broken sheet as failed, with the schema issue as the reason', () => {
            const schemaResult = {
                ok: false,
                missingSheets: [],
                unexpectedSheets: [],
                sheetResults: [{ sheet: 'Customers', ok: false, missingColumns: ['Balance'], unexpectedColumns: [] }]
            };
            const dataTypeResult = { sheetResults: [{ sheet: 'Customers', totalRows: 2, validRows: 2, invalidRows: 0, rowResults: [] }] };

            const report = buildReport({ schemaResult, dataTypeResult });
            expect(report.totalRecords).toBe(2);
            expect(report.passed).toBe(0);
            expect(report.failed).toBe(2);
            expect(report.failures).toEqual([
                expect.objectContaining({ sheet: 'Customers', row: null, reason: expect.stringContaining('Missing required column: "Balance"') })
            ]);
        });

        it('collects per-row, per-field data-type failures with reasons', () => {
            const schemaResult = { ok: true, missingSheets: [], unexpectedSheets: [], sheetResults: [{ sheet: 'Customers', ok: true, missingColumns: [], unexpectedColumns: [] }] };
            const dataTypeResult = {
                sheetResults: [{
                    sheet: 'Customers',
                    totalRows: 2,
                    validRows: 1,
                    invalidRows: 1,
                    rowResults: [
                        { rowNumber: 2, ok: true, errors: [] },
                        { rowNumber: 3, ok: false, errors: [{ field: 'balance', header: 'Balance', reason: 'Expected a number, got "oops"', value: 'oops' }] }
                    ]
                }]
            };

            const report = buildReport({ schemaResult, dataTypeResult });
            expect(report.totalRecords).toBe(2);
            expect(report.passed).toBe(1);
            expect(report.failed).toBe(1);
            expect(report.passRate).toBe(0.5);
            expect(report.failures).toEqual([
                { sheet: 'Customers', row: 3, field: 'Balance', reason: 'Expected a number, got "oops"', value: 'oops' }
            ]);
        });
    });

    describe('reportToWorkbookBuffer', () => {
        it('renders a Summary + Failures workbook that can be read back with exceljs', async () => {
            const ExcelJS = require('exceljs');
            const report = buildReport({
                schemaResult: { ok: true, missingSheets: [], unexpectedSheets: [], sheetResults: [{ sheet: 'Customers', ok: true, missingColumns: [], unexpectedColumns: [] }] },
                dataTypeResult: { sheetResults: [{ sheet: 'Customers', totalRows: 1, validRows: 0, invalidRows: 1, rowResults: [{ rowNumber: 2, ok: false, errors: [{ field: 'id', header: 'ID', reason: 'Required field is empty', value: null }] }] }] }
            });

            const buffer = await reportToWorkbookBuffer(report);
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buffer);

            expect(wb.getWorksheet('Summary')).toBeTruthy();
            expect(wb.getWorksheet('Failures')).toBeTruthy();
            expect(wb.getWorksheet('Failures').rowCount).toBe(2); // header + 1 failure
        });
    });
});
