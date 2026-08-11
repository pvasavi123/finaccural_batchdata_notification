const ExcelJS = require('exceljs');
const ExcelParser = require('../../../../src/modules/excelValidation/parser');

async function buildWorkbookBuffer() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Customers');
    ws.addRow(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
    ws.addRow(['C1', 'Acme Co', 'Acme', 'ap@acme.com', 150.5]);
    ws.addRow(['C2', 'Globex', 'Globex Inc', 'ap@globex.com', 0]);
    ws.addRow([null, null, null, null, null]); // trailing empty row should be skipped
    return wb.xlsx.writeBuffer();
}

describe('ExcelParser', () => {
    it('parses headers and rows from a real .xlsx buffer', async () => {
        const buffer = await buildWorkbookBuffer();
        const parsed = await ExcelParser.parseWorkbook(Buffer.from(buffer));

        expect(parsed.sheetNames).toEqual(['Customers']);
        expect(parsed.sheets.Customers.headers).toEqual(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
        expect(parsed.sheets.Customers.rows).toHaveLength(2);
        expect(parsed.sheets.Customers.rows[0].values).toEqual({
            ID: 'C1', Name: 'Acme Co', 'Company Name': 'Acme', Email: 'ap@acme.com', Balance: 150.5
        });
    });

    it('rejects an empty buffer', async () => {
        await expect(ExcelParser.parseWorkbook(Buffer.alloc(0))).rejects.toThrow();
    });

    it('rejects a buffer that is not a valid xlsx file', async () => {
        await expect(ExcelParser.parseWorkbook(Buffer.from('not an excel file'))).rejects.toThrow();
    });
});
