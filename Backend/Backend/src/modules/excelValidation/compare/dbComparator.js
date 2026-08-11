'use strict';

const { diffRecordSets } = require('./diffEngine');
const { getSchema } = require('../schemas/masterDataSchemas');

/**
 * Excel vs Database Reconciliation
 * -----------------------------------------------------------------
 * Compares parsed "Connections" Excel rows (an ops/finance-maintained
 * roster of which client companies should be connected) against the
 * actual QuickBooksToken / XeroToken rows in the database.
 *
 * Scoped to the authenticated user's own `mail` the same way every
 * other connections endpoint in this app is (see
 * modules/quickbooks/controller.js#listConnections,
 * routes/index.js `/connections`) — this deliberately does NOT accept
 * an arbitrary Sequelize model name from the client. Exposing a
 * "reconcile against model X" endpoint driven by client input would
 * let any authenticated user probe/dump rows from tables (e.g. `User`,
 * `Admin`) that have nothing to do with their own data, which is
 * exactly the kind of IDOR the rest of this codebase is careful to
 * avoid (see the ownership-scoping comments throughout routes/index.js).
 * -----------------------------------------------------------------
 */

const CONNECTIONS_SCHEMA = getSchema('Connections');

/**
 * Fetch this user's QuickBooksToken + XeroToken rows, normalized to the
 * Connections schema's field keys.
 * @param {string} mail - authenticated user's email
 * @returns {Promise<Array<{id:string, companyName:string, email:string, status:string, platform:string}>>}
 */
async function fetchConnectionRows(mail) {
    const { QuickBooksToken, XeroToken } = require('../../../core/database');

    const [qbRows, xeroRows] = await Promise.all([
        QuickBooksToken.findAll({ where: { mail }, raw: true }).catch(() => []),
        XeroToken.findAll({ where: { mail }, raw: true }).catch(() => [])
    ]);

    const normalize = (row, idField, platform) => ({
        id:          row[idField] || '',
        companyName: row.company_name || '',
        email:       row.mail || '',
        status:      row.status || '',
        platform
    });

    return [
        ...qbRows.map((r) => normalize(r, 'realm_id', 'quickbooks')),
        ...xeroRows.map((r) => normalize(r, 'tenant_id', 'xero'))
    ];
}

/**
 * Normalize parsed "Connections" Excel rows into schema-key-keyed
 * objects, matching fetchConnectionRows()'s shape.
 * @param {Array<{values:Object}>} rows - ExcelParser worksheet rows
 * @returns {Array<object>}
 */
function normalizeExcelRows(rows) {
    return rows.map((row) => {
        const normalized = {};
        CONNECTIONS_SCHEMA.columns.forEach((col) => {
            let value = row.values[col.header] ?? null;
            if (col.key === 'platform' && typeof value === 'string') value = value.trim().toLowerCase();
            normalized[col.key] = value;
        });
        return normalized;
    });
}

/**
 * @param {object} opts
 * @param {{headers:string[], rows:Array}} opts.parsedSheet - ExcelParser output for the "Connections" sheet
 * @param {string} opts.mail - authenticated user's email
 * @returns {Promise<{ keyField:string, diff:ReturnType<typeof diffRecordSets> }>}
 */
async function compareConnectionsWithDatabase({ parsedSheet, mail }) {
    const excelRows = normalizeExcelRows(parsedSheet.rows);
    const dbRows = await fetchConnectionRows(mail);

    const fields = CONNECTIONS_SCHEMA.columns.map((c) => c.key).filter((k) => k !== 'id');
    const diff = diffRecordSets({ left: excelRows, right: dbRows, keyField: 'id', fields });

    return { keyField: 'id', diff };
}

module.exports = { fetchConnectionRows, normalizeExcelRows, compareConnectionsWithDatabase, CONNECTIONS_SCHEMA };
