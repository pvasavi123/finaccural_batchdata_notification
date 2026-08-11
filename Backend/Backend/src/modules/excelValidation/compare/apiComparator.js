'use strict';

const { diffRecordSets } = require('./diffEngine');
const { getSchema, API_COMPARABLE_SHEETS } = require('../schemas/masterDataSchemas');
const { ValidationError } = require('../../../core/errors/AppError');

/**
 * Excel vs API Validation
 * -----------------------------------------------------------------
 * Compares parsed Excel rows for one sheet against the SAME data
 * pulled live from QuickBooks/Xero right now (via the existing
 * QuickBooksService / XeroService — no new API integration code, this
 * reuses exactly what exportMasterData() already calls). Surfaces
 * whichever the Excel file has drifted from: values that changed since
 * the export, rows deleted at the source, or rows added since.
 *
 * Field comparison is deliberately loose about *which* fields two
 * platforms both expose (see diffEngine.diffRecordSets — a field only
 * present on one side is skipped, not treated as a mismatch), because
 * QuickBooks and Xero DTOs don't carry identical field sets (e.g. Xero
 * contacts have no `balance`).
 * -----------------------------------------------------------------
 */

/**
 * Fetch the current live rows for one sheet/platform, normalized to the
 * same flat shape the Excel schema's column `key`s expect.
 * @param {string} sheetName - one of API_COMPARABLE_SHEETS
 * @param {'quickbooks'|'xero'} platform
 * @param {string} mail - authenticated user's email (never client-suppliable)
 * @returns {Promise<Array<object>>}
 */
async function fetchLiveRows(sheetName, platform, mail) {
    if (!API_COMPARABLE_SHEETS.includes(sheetName)) {
        throw new ValidationError(`"${sheetName}" cannot be compared against a live API — it has no API source.`);
    }

    if (platform === 'quickbooks') {
        const QuickBooksService = require('../../quickbooks/service');
        switch (sheetName) {
            case 'Company': {
                const company = await QuickBooksService.getCompanyInfo(undefined, mail).catch(() => null);
                return company ? [company] : [];
            }
            case 'Customers': return QuickBooksService.getCustomers(mail);
            case 'Vendors':   return QuickBooksService.getVendors(mail);
            case 'Accounts': {
                const accounts = await QuickBooksService.getAccounts(mail);
                return accounts.map((a) => ({ ...a, balance: a.currentBalance }));
            }
            case 'Classes':   return QuickBooksService.getClasses(mail);
            case 'Locations': return QuickBooksService.getLocations(mail);
            default: return [];
        }
    }

    if (platform === 'xero') {
        const XeroService = require('../../xero/service');
        switch (sheetName) {
            case 'Company': {
                const org = await XeroService.getOrganisation(mail).catch(() => null);
                return org ? [org] : [];
            }
            case 'Customers': {
                const contacts = await XeroService.getContacts(mail);
                return contacts.filter((c) => c.isCustomer);
            }
            case 'Vendors': {
                const contacts = await XeroService.getContacts(mail);
                return contacts.filter((c) => c.isSupplier);
            }
            case 'Accounts':  return XeroService.getAccounts(mail);
            case 'Classes':   return XeroService.getClasses(mail);
            case 'Locations': return XeroService.getLocations(mail);
            default: return [];
        }
    }

    throw new ValidationError(`Unsupported platform "${platform}". Expected "quickbooks" or "xero".`);
}

/**
 * Normalize parsed Excel rows (header-keyed) into schema-key-keyed
 * objects, e.g. { 'Company Name': 'Acme' } -> { companyName: 'Acme' }.
 * @param {Array<{values:Object}>} rows - ExcelParser worksheet rows
 * @param {{columns:Array<{key:string, header:string}>}} schema
 * @returns {Array<object>}
 */
function normalizeExcelRows(rows, schema) {
    return rows.map((row) => {
        const normalized = {};
        schema.columns.forEach((col) => {
            normalized[col.key] = row.values[col.header] ?? null;
        });
        return normalized;
    });
}

/**
 * @param {object} opts
 * @param {string} opts.sheetName
 * @param {{headers:string[], rows:Array}} opts.parsedSheet - ExcelParser output for this one sheet
 * @param {'quickbooks'|'xero'} opts.platform
 * @param {string} opts.mail
 * @returns {Promise<{ sheet:string, keyField:string, diff:ReturnType<typeof diffRecordSets> }>}
 */
async function compareSheetWithApi({ sheetName, parsedSheet, platform, mail }) {
    const schema = getSchema(sheetName);
    const excelRows = normalizeExcelRows(parsedSheet.rows, schema);
    const liveRows = await fetchLiveRows(sheetName, platform, mail);

    const fields = schema.columns.map((c) => c.key).filter((k) => k !== 'id');
    const diff = diffRecordSets({ left: excelRows, right: liveRows, keyField: 'id', fields });

    return { sheet: sheetName, keyField: 'id', diff };
}

module.exports = { fetchLiveRows, normalizeExcelRows, compareSheetWithApi };
