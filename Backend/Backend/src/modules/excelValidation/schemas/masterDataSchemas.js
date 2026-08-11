'use strict';

/**
 * Master Data Excel Schemas
 * -----------------------------------------------------------------
 * Excel Schema Validation / Excel Data-Type Validation.
 *
 * Column-level contracts for every worksheet FinAccrual actually
 * produces via QuickbooksController#exportMasterData /
 * XeroController#exportMasterData ("quickbooks_master_data.xlsx" /
 * "xero_master_data.xlsx") — same worksheet names, same header order,
 * same source DTOs (see modules/quickbooks/mapper.js and
 * modules/xero/mapper.js). Reusing those exact headers means a file a
 * user actually exported from this app will validate cleanly out of
 * the box; anything that drifts from them (hand-edited, exported from
 * elsewhere, stale template) is exactly what this module should catch.
 *
 * `key` is the normalized, camelCase field name used everywhere
 * downstream of parsing (data-type errors, API/DB comparisons) so the
 * rest of the module never has to deal with spaced, human header text.
 * `header` is the literal first-row Excel column title we expect.
 * -----------------------------------------------------------------
 */

const TYPES = Object.freeze({
    STRING:  'string',
    NUMBER:  'number',
    BOOLEAN: 'boolean',
    EMAIL:   'email',
    DATE:    'date',
    ENUM:    'enum'
});

const ACTIVE_STATUS_VALUES = ['Active', 'Inactive'];
const CONNECTION_STATUS_VALUES = ['Not Synced', 'Active', 'Disconnected'];
const PLATFORM_VALUES = ['quickbooks', 'xero'];

const SHEET_SCHEMAS = {
    Company: {
        sheetName: 'Company',
        columns: [
            { key: 'id',        header: 'ID',           type: TYPES.STRING, required: true },
            { key: 'name',      header: 'Company Name', type: TYPES.STRING, required: true },
            { key: 'legalName', header: 'Legal Name',   type: TYPES.STRING, required: false }
        ]
    },

    Customers: {
        sheetName: 'Customers',
        columns: [
            { key: 'id',          header: 'ID',           type: TYPES.STRING, required: true },
            { key: 'name',        header: 'Name',         type: TYPES.STRING, required: true },
            { key: 'companyName', header: 'Company Name', type: TYPES.STRING, required: false },
            { key: 'email',       header: 'Email',        type: TYPES.EMAIL,  required: false },
            { key: 'balance',     header: 'Balance',      type: TYPES.NUMBER, required: true }
        ]
    },

    Vendors: {
        sheetName: 'Vendors',
        columns: [
            { key: 'id',          header: 'ID',           type: TYPES.STRING, required: true },
            { key: 'name',        header: 'Name',         type: TYPES.STRING, required: true },
            { key: 'companyName', header: 'Company Name', type: TYPES.STRING, required: false },
            { key: 'email',       header: 'Email',        type: TYPES.EMAIL,  required: false },
            { key: 'balance',     header: 'Balance',      type: TYPES.NUMBER, required: true }
        ]
    },

    Accounts: {
        sheetName: 'Accounts',
        columns: [
            { key: 'id',             header: 'ID',           type: TYPES.STRING, required: true },
            { key: 'acctNum',        header: 'Acct #',       type: TYPES.STRING, required: false },
            { key: 'name',           header: 'Name',         type: TYPES.STRING, required: true },
            { key: 'accountType',    header: 'Account Type', type: TYPES.STRING, required: true },
            { key: 'accountSubType', header: 'Sub Type',     type: TYPES.STRING, required: false },
            { key: 'balance',        header: 'Balance',      type: TYPES.NUMBER, required: true }
        ]
    },

    Classes: {
        sheetName: 'Classes',
        columns: [
            { key: 'id',     header: 'ID',     type: TYPES.STRING, required: true },
            { key: 'name',   header: 'Name',   type: TYPES.STRING, required: true },
            { key: 'status', header: 'Status', type: TYPES.ENUM,   required: true, enumValues: ACTIVE_STATUS_VALUES }
        ]
    },

    Locations: {
        sheetName: 'Locations',
        columns: [
            { key: 'id',     header: 'ID',     type: TYPES.STRING, required: true },
            { key: 'name',   header: 'Name',   type: TYPES.STRING, required: true },
            { key: 'status', header: 'Status', type: TYPES.ENUM,   required: true, enumValues: ACTIVE_STATUS_VALUES }
        ]
    },

    // Excel vs Database Reconciliation target — mirrors the columns an
    // ops/finance user would maintain in a roster spreadsheet of "which
    // client companies should be connected", reconciled against the
    // live QuickBooksToken / XeroToken rows (core/database).
    Connections: {
        sheetName: 'Connections',
        columns: [
            { key: 'id',          header: 'Company ID',   type: TYPES.STRING, required: true },
            { key: 'companyName', header: 'Company Name', type: TYPES.STRING, required: true },
            { key: 'email',       header: 'Owner Email',  type: TYPES.EMAIL,  required: true },
            { key: 'status',      header: 'Status',       type: TYPES.ENUM,   required: true, enumValues: CONNECTION_STATUS_VALUES },
            { key: 'platform',    header: 'Platform',     type: TYPES.ENUM,   required: true, enumValues: PLATFORM_VALUES }
        ]
    }
};

// The full multi-sheet workbook produced by /api/quickbooks/export and
// /api/xero/export — every one of these sheets is expected to be
// present (Company is the one exception: it's only added when the live
// company lookup succeeded, see exportMasterData).
const MASTER_DATA_WORKBOOK_SHEETS = ['Company', 'Customers', 'Vendors', 'Accounts', 'Classes', 'Locations'];
const OPTIONAL_WORKBOOK_SHEETS = ['Company'];

// Sheets whose rows can be diffed against a live QuickBooks/Xero API
// pull (Excel vs API Validation). Connections is DB-only.
const API_COMPARABLE_SHEETS = ['Company', 'Customers', 'Vendors', 'Accounts', 'Classes', 'Locations'];

function getSchema(sheetName) {
    const schema = SHEET_SCHEMAS[sheetName];
    if (!schema) {
        const known = Object.keys(SHEET_SCHEMAS).join(', ');
        throw new Error(`Unknown Excel schema "${sheetName}". Known schemas: ${known}`);
    }
    return schema;
}

module.exports = {
    TYPES,
    ACTIVE_STATUS_VALUES,
    CONNECTION_STATUS_VALUES,
    PLATFORM_VALUES,
    SHEET_SCHEMAS,
    MASTER_DATA_WORKBOOK_SHEETS,
    OPTIONAL_WORKBOOK_SHEETS,
    API_COMPARABLE_SHEETS,
    getSchema
};
