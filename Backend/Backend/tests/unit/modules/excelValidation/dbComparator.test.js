const mockQbFindAll = jest.fn();
const mockXeroFindAll = jest.fn();

jest.mock('../../../../src/core/database', () => ({
    QuickBooksToken: { findAll: (...args) => mockQbFindAll(...args) },
    XeroToken: { findAll: (...args) => mockXeroFindAll(...args) }
}));

const { fetchConnectionRows, normalizeExcelRows, compareConnectionsWithDatabase } = require('../../../../src/modules/excelValidation/compare/dbComparator');

describe('dbComparator', () => {
    afterEach(() => jest.clearAllMocks());

    describe('fetchConnectionRows', () => {
        it('normalizes QuickBooksToken + XeroToken rows into the Connections schema shape, scoped by mail', async () => {
            mockQbFindAll.mockResolvedValue([{ realm_id: 'QB1', company_name: 'Acme', mail: 'user@example.com', status: 'Active' }]);
            mockXeroFindAll.mockResolvedValue([{ tenant_id: 'XR1', company_name: 'Globex', mail: 'user@example.com', status: 'Not Synced' }]);

            const rows = await fetchConnectionRows('user@example.com');

            expect(mockQbFindAll).toHaveBeenCalledWith({ where: { mail: 'user@example.com' }, raw: true });
            expect(mockXeroFindAll).toHaveBeenCalledWith({ where: { mail: 'user@example.com' }, raw: true });
            expect(rows).toEqual([
                { id: 'QB1', companyName: 'Acme', email: 'user@example.com', status: 'Active', platform: 'quickbooks' },
                { id: 'XR1', companyName: 'Globex', email: 'user@example.com', status: 'Not Synced', platform: 'xero' }
            ]);
        });
    });

    describe('normalizeExcelRows', () => {
        it('lowercases the platform column for comparison', () => {
            const rows = normalizeExcelRows([
                { values: { 'Company ID': 'QB1', 'Company Name': 'Acme', 'Owner Email': 'user@example.com', Status: 'Active', Platform: 'QuickBooks' } }
            ]);
            expect(rows).toEqual([{ id: 'QB1', companyName: 'Acme', email: 'user@example.com', status: 'Active', platform: 'quickbooks' }]);
        });
    });

    describe('compareConnectionsWithDatabase', () => {
        it('flags a status drift between Excel and the DB', async () => {
            mockQbFindAll.mockResolvedValue([{ realm_id: 'QB1', company_name: 'Acme', mail: 'user@example.com', status: 'Disconnected' }]);
            mockXeroFindAll.mockResolvedValue([]);

            const parsedSheet = {
                headers: ['Company ID', 'Company Name', 'Owner Email', 'Status', 'Platform'],
                rows: [{ __rowNumber: 2, values: { 'Company ID': 'QB1', 'Company Name': 'Acme', 'Owner Email': 'user@example.com', Status: 'Active', Platform: 'quickbooks' } }]
            };

            const { diff } = await compareConnectionsWithDatabase({ parsedSheet, mail: 'user@example.com' });
            expect(diff.mismatchedCount).toBe(1);
            expect(diff.mismatched[0].differences).toEqual([{ field: 'status', leftValue: 'Active', rightValue: 'Disconnected' }]);
        });
    });
});
