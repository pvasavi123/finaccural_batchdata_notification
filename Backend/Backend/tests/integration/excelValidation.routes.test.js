const request = require('supertest');

// Fake an authenticated request without needing a real JWT/DB user —
// same approach as mocking any other dependency layer in this test
// suite (see tests/integration/admin.routes.test.js mocking the
// service layer). req.user.email is what every excel-validation
// controller method scopes its API/DB lookups by.
jest.mock('../../src/modules/auth/auth.middleware', () => ({
    authenticate: (req, res, next) => {
        req.user = { userId: 1, email: 'user@example.com', role: 'user', name: 'Test User' };
        next();
    }
}));

jest.mock('../../src/modules/excelValidation/service');

const app = require('../../src/app');
const ExcelValidationService = require('../../src/modules/excelValidation/service');

const SOME_BASE64 = Buffer.from('placeholder-not-a-real-xlsx').toString('base64');

describe('Excel Validation Routes Integration', () => {
    afterEach(() => jest.clearAllMocks());

    describe('POST /api/excel-validation/schema-check', () => {
        it('returns 400 when fileBase64 is missing', async () => {
            const res = await request(app)
                .post('/api/excel-validation/schema-check')
                .send({ schema: 'Customers' });

            expect(res.status).toBe(400);
        });

        it('returns 400 for an unknown schema name', async () => {
            const res = await request(app)
                .post('/api/excel-validation/schema-check')
                .send({ fileBase64: SOME_BASE64, schema: 'NotARealSheet' });

            expect(res.status).toBe(400);
        });

        it('returns 200 with the schema validation result on success', async () => {
            ExcelValidationService.validateSchema.mockResolvedValue({
                parsed: {},
                schemaResult: { ok: true, missingSheets: [], unexpectedSheets: [], sheetResults: [] }
            });

            const res = await request(app)
                .post('/api/excel-validation/schema-check')
                .send({ fileBase64: SOME_BASE64, schema: 'Customers' });

            expect(res.status).toBe(200);
            expect(res.body.ok).toBe(true);
            expect(res.body.schema).toBe('Customers');
        });
    });

    describe('POST /api/excel-validation/data-type-check', () => {
        it('returns 200 with the data-type validation result', async () => {
            ExcelValidationService.validateDataTypes.mockResolvedValue({
                parsed: {},
                dataTypeResult: { sheetResults: [{ sheet: 'Customers', totalRows: 1, validRows: 1, invalidRows: 0, rowResults: [] }] }
            });

            const res = await request(app)
                .post('/api/excel-validation/data-type-check')
                .send({ fileBase64: SOME_BASE64, schema: 'Customers' });

            expect(res.status).toBe(200);
            expect(res.body.sheetResults[0].sheet).toBe('Customers');
        });
    });

    describe('POST /api/excel-validation/report', () => {
        it('returns the JSON PASS/FAIL report by default', async () => {
            ExcelValidationService.generateReport.mockResolvedValue({
                parsed: {}, schemaResult: {}, dataTypeResult: {},
                report: { generatedAt: '2026-08-11T00:00:00.000Z', totalRecords: 5, passed: 4, failed: 1, passRate: 0.8, schemaOk: true, missingSheets: [], unexpectedSheets: [], bySheet: [], failures: [] }
            });

            const res = await request(app)
                .post('/api/excel-validation/report')
                .send({ fileBase64: SOME_BASE64, schema: 'MasterData' });

            expect(res.status).toBe(200);
            expect(res.body.totalRecords).toBe(5);
            expect(res.body.passed).toBe(4);
            expect(res.body.failed).toBe(1);
        });

        it('streams an .xlsx workbook when ?format=xlsx is requested', async () => {
            ExcelValidationService.generateReportWorkbook.mockResolvedValue({
                report: { passed: 4, failed: 1 },
                workbookBuffer: Buffer.from('fake-xlsx-bytes')
            });

            const res = await request(app)
                .post('/api/excel-validation/report?format=xlsx')
                .send({ fileBase64: SOME_BASE64, schema: 'MasterData' });

            expect(res.status).toBe(200);
            expect(res.headers['content-type']).toContain('spreadsheetml.sheet');
            expect(res.headers['content-disposition']).toContain('excel_validation_report.xlsx');
            expect(res.headers['x-report-passed']).toBe('4');
        });
    });

    describe('POST /api/excel-validation/vs-api', () => {
        it('rejects an unsupported platform before hitting the service', async () => {
            const res = await request(app)
                .post('/api/excel-validation/vs-api')
                .send({ fileBase64: SOME_BASE64, sheet: 'Customers', platform: 'sap' });

            expect(res.status).toBe(400);
            expect(ExcelValidationService.compareWithApi).not.toHaveBeenCalled();
        });

        it('scopes the comparison to the authenticated user, never a client-supplied identity', async () => {
            ExcelValidationService.compareWithApi.mockResolvedValue({
                parsed: {}, keyField: 'id',
                diff: { totalLeft: 1, totalRight: 1, matchedCount: 1, mismatchedCount: 0, missingInRightCount: 0, missingInLeftCount: 0, matched: [{ key: '1' }], mismatched: [], missingInRight: [], missingInLeft: [] }
            });

            const res = await request(app)
                .post('/api/excel-validation/vs-api')
                .send({ fileBase64: SOME_BASE64, sheet: 'Customers', platform: 'quickbooks', mail: 'attacker@example.com' });

            expect(res.status).toBe(200);
            expect(ExcelValidationService.compareWithApi).toHaveBeenCalledWith(
                expect.any(Buffer), 'Customers', 'quickbooks', 'user@example.com'
            );
        });
    });

    describe('POST /api/excel-validation/vs-database', () => {
        it('returns the reconciliation diff', async () => {
            ExcelValidationService.compareWithDatabase.mockResolvedValue({
                parsed: {}, keyField: 'id',
                diff: { totalLeft: 1, totalRight: 1, matchedCount: 0, mismatchedCount: 1, missingInRightCount: 0, missingInLeftCount: 0, matched: [], mismatched: [{ key: 'QB1', differences: [{ field: 'status', leftValue: 'Active', rightValue: 'Disconnected' }] }], missingInRight: [], missingInLeft: [] }
            });

            const res = await request(app)
                .post('/api/excel-validation/vs-database')
                .send({ fileBase64: SOME_BASE64 });

            expect(res.status).toBe(200);
            expect(res.body.sheet).toBe('Connections');
            expect(res.body.mismatchedCount).toBe(1);
        });
    });

    afterAll(async () => {
        const { sequelize } = require('../../src/core/database');
        await sequelize.close();
    });
});
