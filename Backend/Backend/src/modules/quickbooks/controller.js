'use strict';

const querystring  = require('querystring');
const exceljs      = require('exceljs');
const config       = require('../../core/config');
const CONSTANTS    = require('../../core/constants');
const { generateOAuthState } = require('../../core/helpers');
const QuickBooksService    = require('./service');
const QuickBooksTokenRepository = require('./repository');
const { ValidationError } = require('../../core/errors/AppError');

/**
 * QuickbooksController
 * -----------------------------------------------------------------
 * Handles all incoming HTTP requests for the QuickBooks module.
 * Delegates all business logic to QuickBooksService.
 * Does NOT contain any data-transformation or mapping logic —
 * that responsibility lives in mapper.js (used by the Service).
 * -----------------------------------------------------------------
 */
class QuickbooksController {

    /**
     * GET /api/quickbooks/connect
     * Generates the QuickBooks OAuth authorization URL and redirects.
     */
    connectQuickbooks = async (req, res, next) => {
        try {
            const { QuickBooksToken } = require('../../core/database');
            const mail = req.query.mail || req.session?.user_mail || req.session?.admin?.email || req.session?.googleUser?.email || null;
            const { Op } = require('sequelize');
            const whereClause = { status: { [Op.ne]: 'Disconnected' } };
            if (mail) whereClause.mail = mail;
            const qbCount = await QuickBooksToken.count({ where: whereClause });
            const tier = (req.query.tier || 'pro').toLowerCase();

            let maxAllowed = 10;
            if (tier === 'basic') maxAllowed = 1;
            else if (tier === 'standard') maxAllowed = 3;

            if (qbCount >= maxAllowed) {
                return res.send(`
                    <html>
                        <body style="font-family:sans-serif; text-align:center; padding: 40px; background:#fff1f2; color:#9f1239;">
                            <div style="font-size: 50px; margin-bottom: 20px;">⚠️</div>
                            <h2>Connection Limit Reached</h2>
                            <p style="font-size: 14px; color: #4b5563;">Your subscription tier (${tier.toUpperCase()}) allows a maximum of ${maxAllowed} connected company.</p>
                            <p style="font-size: 14px; color: #4b5563;">Please disconnect an existing company or upgrade your plan to connect more.</p>
                            <button onclick="window.close()" style="margin-top: 20px; padding:10px 20px; background:#be123c; color:white; border:none; border-radius:5px; cursor:pointer; font-weight: bold;">Close Window</button>
                        </body>
                    </html>
                `);
            }

            const state = generateOAuthState();
            req.session.oauth_state = state;
            req.session.user_mail = req.query.mail || null;

            const params = {
                client_id:     config.QB.CLIENT_ID,
                response_type: 'code',
                scope:         CONSTANTS.QUICKBOOKS.SCOPES,
                redirect_uri:  config.QB.REDIRECT_URI,
                state
            };

            const authUrl = `${CONSTANTS.QUICKBOOKS.AUTH_URL}?${querystring.stringify(params)}`;
            res.redirect(authUrl);
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/quickbooks/callback
     * Handles the OAuth callback, exchanges code for tokens.
     */
    quickbooksCallback = async (req, res, next) => {
        try {
            const { code, realmId } = req.query;
            const mail = req.session?.user_mail || req.session?.admin?.email || req.session?.googleUser?.email || null;
            const sessionInfo = JSON.stringify(req.session || {});
            await QuickBooksService.exchangeAndSaveToken(code, realmId, sessionInfo, mail);
            return res.send(CONSTANTS.QUICKBOOKS.SUCCESS_HTML);
        } catch (error) {
            const details = JSON.stringify(error.response?.data || error.message);
            next(new ValidationError('Failed to connect QuickBooks. Please try again.', details));
        }
    };

    /**
     * GET /api/quickbooks/tokens
     * Returns all stored QuickBooks OAuth tokens (for debugging).
     */
    listQuickbooksTokens = async (req, res, next) => {
        try {
            const tokens = await QuickBooksTokenRepository.getAllTokens();
            res.json({ tokens });
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/quickbooks/customers
     * Returns a list of mapped CustomerDTOs.
     */
    getCustomers = async (req, res, next) => {
        try {
            const customers = await QuickBooksService.getCustomers();
            res.json({ customers });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/vendors
     * Returns a list of mapped VendorDTOs.
     */
    getVendors = async (req, res, next) => {
        try {
            const vendors = await QuickBooksService.getVendors();
            res.json({ vendors });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/accounts
     * Returns a list of mapped AccountDTOs.
     */
    getAccounts = async (req, res, next) => {
        try {
            const accounts = await QuickBooksService.getAccounts();
            res.json({ accounts });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/classes
     * Returns a list of mapped ClassDTOs.
     */
    getClasses = async (req, res, next) => {
        try {
            const classes = await QuickBooksService.getClasses();
            res.json({ classes });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/locations
     * Returns a list of mapped LocationDTOs.
     */
    getLocations = async (req, res, next) => {
        try {
            const locations = await QuickBooksService.getLocations();
            res.json({ locations });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/company
     * Returns company information DTO.
     */
    getCompanyInfo = async (req, res, next) => {
        try {
            const company = await QuickBooksService.getCompanyInfo();
            res.json({ company });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/export
     * Exports company, customers, vendors, accounts, classes, and locations as an Excel file.
     */
    exportMasterData = async (req, res, next) => {
        try {
            const [company, customers, vendors, accounts, classes, locations] = await Promise.all([
                QuickBooksService.getCompanyInfo().catch(() => null),
                QuickBooksService.getCustomers().catch(() => []),
                QuickBooksService.getVendors().catch(() => []),
                QuickBooksService.getAccounts().catch(() => []),
                QuickBooksService.getClasses().catch(() => []),
                QuickBooksService.getLocations().catch(() => [])
            ]);

            const wb = new exceljs.Workbook();

            if (company) {
                const wsCompany = wb.addWorksheet('Company');
                wsCompany.addRow(['ID', 'Company Name', 'Legal Name']);
                wsCompany.addRow([company.id, company.name, company.legalName]);
            }

            const wsCustomers = wb.addWorksheet('Customers');
            wsCustomers.addRow(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
            wsCustomers.addRows(customers.map(c => [c.id, c.name, c.companyName, c.email, c.balance]));

            const wsVendors = wb.addWorksheet('Vendors');
            wsVendors.addRow(['ID', 'Name', 'Company Name', 'Email', 'Balance']);
            wsVendors.addRows(vendors.map(v => [v.id, v.name, v.companyName, v.email, v.balance]));

            const wsAccounts = wb.addWorksheet('Accounts');
            wsAccounts.addRow(['ID', 'Acct #', 'Name', 'Account Type', 'Sub Type', 'Balance']);
            wsAccounts.addRows(accounts.map(a => [a.id, a.acctNum, a.name, a.accountType, a.accountSubType, a.currentBalance]));

            const wsClasses = wb.addWorksheet('Classes');
            wsClasses.addRow(['ID', 'Name', 'Status']);
            wsClasses.addRows(classes.map(c => [c.id, c.name, c.active]));

            const wsLocations = wb.addWorksheet('Locations');
            wsLocations.addRow(['ID', 'Name', 'Status']);
            wsLocations.addRows(locations.map(l => [l.id, l.name, l.active]));

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', 'attachment; filename="quickbooks_master_data.xlsx"');

            await wb.xlsx.write(res);
            res.end();
        } catch (error) {
            next(error);
        }
    };

    /**
     * POST /api/quickbooks/disconnect
     * Clears all stored QuickBooks tokens.
     */
    disconnectQuickbooks = async (req, res, next) => {
        try {
            await QuickBooksTokenRepository.clearTokens();
            res.json({ success: true, message: 'QuickBooks tokens cleared successfully.' });
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/quickbooks/connections
     */
    listConnections = async (req, res, next) => {
        try {
            const mail = req.query.mail
                || req.session?.user_mail
                || req.session?.admin?.email
                || req.session?.googleUser?.email
                || null;

            const list = await QuickBooksService.listConnections(mail);
            return res.json(list);
        } catch (err) {
            return next(err);
        }
    };

    /**
     * GET /api/quickbooks/connections/stats
     */
    getConnectionStats = async (req, res, next) => {
        try {
            const mail = req.query.mail || req.session?.user_mail || null;
            const plan = req.query.plan || 'pro';

            const stats = {
                plan: plan.toLowerCase(),
                maxPerPlatform: 10,
                quickbooks: { connected: 0, remaining: 10 }
            };

            if (plan === 'basic')    stats.maxPerPlatform = 1;
            else if (plan === 'standard') stats.maxPerPlatform = 3;

            const qbStats = await QuickBooksService.getConnectionStats(mail, plan);
            stats.quickbooks = {
                connected: qbStats.connected,
                remaining: qbStats.remaining
            };

            return res.json(stats);
        } catch (err) {
            return next(err);
        }
    };

    /**
     * DELETE /api/quickbooks/connections/:id
     */
    disconnectConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const success = await QuickBooksService.disconnectConnection(companyId);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * POST /api/quickbooks/connections/:id/activate
     */
    activateConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const success = await QuickBooksService.activateConnection(companyId);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * PATCH /api/quickbooks/connections/:id/rename
     */
    renameConnection = async (req, res, next) => {
        try {
            const companyId = req.params.id;
            const { companyName } = req.body;
            if (!companyName) {
                throw new ValidationError('companyName is required.');
            }

            const success = await QuickBooksService.renameConnection(companyId, companyName);
            return res.json({ success: !!success });
        } catch (err) {
            return next(err);
        }
    };

    /**
     * GET /api/quickbooks/pull-master-data?companyId=...&tier=...
     */
    pullMasterData = async (req, res, next) => {
        try {
            const { companyId, tier } = req.query;
            
            const aggregated = await QuickBooksService.pullMasterData(companyId, tier);

            if (!aggregated) {
                const { AppError } = require('../../core/errors/AppError');
                throw new AppError('The requested resource was not found.', 404, 'ERR_NOT_FOUND', `No active connections found for quickbooks.`);
            }

            return res.json({
                company:   aggregated.company.length === 1 ? aggregated.company[0] : aggregated.company,
                customers: aggregated.customers,
                vendors:   aggregated.vendors,
                accounts:  aggregated.accounts,
                classes:   aggregated.classes,
                locations: aggregated.locations
            });
        } catch (err) {
            return next(err);
        }
    };
}

module.exports = new QuickbooksController();
