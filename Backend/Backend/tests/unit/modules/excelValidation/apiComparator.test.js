jest.mock('../../../../src/modules/quickbooks/service');
jest.mock('../../../../src/modules/xero/service');

const QuickBooksService = require('../../../../src/modules/quickbooks/service');
const XeroService = require('../../../../src/modules/xero/service');
const { fetchLiveRows, normalizeExcelRows, compareSheetWithApi } = require('../../../../src/modules/excelValidation/compare/apiComparator');
const { getSchema } = require('../../../../src/modules/excelValidation/schemas/masterDataSchemas');

describe('apiComparator', () => {
    afterEach(() => jest.clearAllMocks());

    describe('fetchLiveRows', () => {
        it('pulls QuickBooks customers as-is', async () => {
            QuickBooksService.getCustomers.mockResolvedValue([{ id: '1', name: 'Acme' }]);
            const rows = await fetchLiveRows('Customers', 'quickbooks', 'user@example.com');
            expect(rows).toEqual([{ id: '1', name: 'Acme' }]);
            expect(QuickBooksService.getCustomers).toHaveBeenCalledWith('user@example.com');
        });

        it('maps QuickBooks Accounts.currentBalance onto a balance field', async () => {
            QuickBooksService.getAccounts.mockResolvedValue([{ id: '1', name: 'Checking', currentBalance: 500 }]);
            const rows = await fetchLiveRows('Accounts', 'quickbooks', 'user@example.com');
            expect(rows).toEqual([{ id: '1', name: 'Checking', currentBalance: 500, balance: 500 }]);
        });

        it('splits Xero contacts into customers vs vendors via isCustomer/isSupplier', async () => {
            XeroService.getContacts.mockResolvedValue([
                { id: '1', name: 'Cust Co', isCustomer: true, isSupplier: false },
                { id: '2', name: 'Vendor Co', isCustomer: false, isSupplier: true }
            ]);

            const customers = await fetchLiveRows('Customers', 'xero', 'user@example.com');
            expect(customers).toEqual([{ id: '1', name: 'Cust Co', isCustomer: true, isSupplier: false }]);

            const vendors = await fetchLiveRows('Vendors', 'xero', 'user@example.com');
            expect(vendors).toEqual([{ id: '2', name: 'Vendor Co', isCustomer: false, isSupplier: true }]);
        });

        it('rejects an unsupported platform', async () => {
            await expect(fetchLiveRows('Customers', 'sap', 'user@example.com')).rejects.toThrow(/Unsupported platform/);
        });

        it('rejects a sheet that has no API source', async () => {
            await expect(fetchLiveRows('Connections', 'quickbooks', 'user@example.com')).rejects.toThrow();
        });
    });

    describe('normalizeExcelRows', () => {
        it('maps header-keyed Excel rows onto schema keys', () => {
            const schema = getSchema('Customers');
            const rows = normalizeExcelRows(
                [{ values: { ID: 'C1', Name: 'Acme', 'Company Name': 'Acme Co', Email: 'a@acme.com', Balance: 100 } }],
                schema
            );
            expect(rows).toEqual([{ id: 'C1', name: 'Acme', companyName: 'Acme Co', email: 'a@acme.com', balance: 100 }]);
        });
    });

    describe('compareSheetWithApi', () => {
        it('flags a balance mismatch between Excel and live QuickBooks data', async () => {
            QuickBooksService.getCustomers.mockResolvedValue([{ id: 'C1', name: 'Acme', companyName: 'Acme Co', email: 'a@acme.com', balance: 999 }]);

            const parsedSheet = {
                headers: ['ID', 'Name', 'Company Name', 'Email', 'Balance'],
                rows: [{ __rowNumber: 2, values: { ID: 'C1', Name: 'Acme', 'Company Name': 'Acme Co', Email: 'a@acme.com', Balance: 100 } }]
            };

            const { diff } = await compareSheetWithApi({ sheetName: 'Customers', parsedSheet, platform: 'quickbooks', mail: 'user@example.com' });
            expect(diff.mismatchedCount).toBe(1);
            expect(diff.mismatched[0].differences).toEqual([{ field: 'balance', leftValue: 100, rightValue: 999 }]);
        });
    });
});
