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
     *
     * Requires authentication (the frontend passes the JWT as ?token=...
     * since this is a browser navigation, not a fetch). The owning email
     * is taken exclusively from the verified token (req.user.email), never
     * from a client-suppliable query param — otherwise anyone could
     * initiate a connect flow tagged with someone else's email and inject
     * a connection into that other user's account.
     */
    connectQuickbooks = async (req, res, next) => {
        try {
            const { QuickBooksToken } = require('../../core/database');
            const mail = req.user.email;
            const { Op } = require('sequelize');
            const whereClause = { status: { [Op.ne]: 'Disconnected' }, mail };
            const qbCount = await QuickBooksToken.count({ where: whereClause });
            const tier = (req.query.tier || 'pro').toLowerCase();

            let maxAllowed = 10;
            if (tier === 'trial') maxAllowed = 1;
            else if (tier === 'basic') maxAllowed = 1;
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
            req.session.user_mail = mail;

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
     * Returns the authenticated user's own stored QuickBooks OAuth tokens.
     */
    listQuickbooksTokens = async (req, res, next) => {
        try {
            const tokens = await QuickBooksTokenRepository.getAllTokens(req.user.email);
            res.json({ tokens });
        } catch (error) {
            next(error);
        }
    };

    /**
     * GET /api/quickbooks/customers
     * Returns a list of mapped CustomerDTOs for the authenticated user's companies.
     */
    getCustomers = async (req, res, next) => {
        try {
            const customers = await QuickBooksService.getCustomers(req.user.email);
            res.json({ customers });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/vendors
     * Returns a list of mapped VendorDTOs for the authenticated user's companies.
     */
    getVendors = async (req, res, next) => {
        try {
            const vendors = await QuickBooksService.getVendors(req.user.email);
            res.json({ vendors });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/accounts
     * Returns a list of mapped AccountDTOs for the authenticated user's companies.
     */
    getAccounts = async (req, res, next) => {
        try {
            const accounts = await QuickBooksService.getAccounts(req.user.email);
            res.json({ accounts });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/classes
     * Returns a list of mapped ClassDTOs for the authenticated user's companies.
     */
    getClasses = async (req, res, next) => {
        try {
            const classes = await QuickBooksService.getClasses(req.user.email);
            res.json({ classes });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/locations
     * Returns a list of mapped LocationDTOs for the authenticated user's companies.
     */
    getLocations = async (req, res, next) => {
        try {
            const locations = await QuickBooksService.getLocations(req.user.email);
            res.json({ locations });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/company
     * Returns company information DTO for the authenticated user's companies.
     */
    getCompanyInfo = async (req, res, next) => {
        try {
            const company = await QuickBooksService.getCompanyInfo(undefined, req.user.email);
            res.json({ company });
        } catch (err) {
            next(err);
        }
    };

    /**
     * GET /api/quickbooks/export
     * Exports company, customers, vendors, accounts, classes, and locations
     * as an Excel file, scoped to the authenticated user's companies.
     *
     * Fetches in 10-record batches per entity instead of one big
     * unpaginated call per entity. Each batch still fires every entity's
     * API concurrently via Promise.all — same shape as the original
     * single-shot Promise.all below, just repeated batch-by-batch — so
     * customers/vendors/accounts/classes/locations all pull records
     * 1–10 together, then 11–20 together, and so on, until every one of
     * them has exhausted every connected company's data. Company info
     * isn't paginated (one record per company) and is fetched once,
     * up front, alongside a token→orgName lookup so the per-batch entity
     * fetches don't each need their own CompanyInfo round trip.
     */
    exportMasterData = async (req, res, next) => {
        try {
            const mail = req.user.email;
            const BATCH_SIZE = 10;

            const { tokens: allTokens, company, orgNameByTokenId } =
                await QuickBooksService.getCompanyInfoAndOrgNames(mail)
                    .catch(() => ({ tokens: [], company: null, orgNameByTokenId: new Map() }));

            // Per-entity list of tokens still known to have more pages —
            // shrinks independently as each token/company reports its
            // last (short) page, so a company with fewer records simply
            // stops being queried for that entity while others with more
            // data keep going. No two entities share a list, since one
            // entity finishing early for a company must not affect the
            // others' pagination.
            let customersTokens = allTokens.slice();
            let vendorsTokens   = allTokens.slice();
            let accountsTokens  = allTokens.slice();
            let classesTokens   = allTokens.slice();
            let locationsTokens = allTokens.slice();

            const customers = [];
            const vendors   = [];
            const accounts  = [];
            const classes   = [];
            const locations = [];

            const dropExhausted = (list, exhaustedIds) =>
                list.filter(t => !exhaustedIds.has(t.companyId || t.realm_id));
            const allTokenIds = (list) => new Set(list.map(t => t.companyId || t.realm_id));
            const emptyPage = () => ({ records: [], exhaustedTokenIds: new Set() });

            let startPosition = 1;
            let batchCount = 0;
            // Safety valve only — each entity's token list can only
            // shrink every iteration, so the loop is guaranteed to end
            // long before this; it just guards against an infinite loop
            // if that invariant is ever broken by a future change.
            const MAX_BATCHES = 100000;

            while (
                (customersTokens.length || vendorsTokens.length || accountsTokens.length ||
                 classesTokens.length || locationsTokens.length) &&
                batchCount < MAX_BATCHES
            ) {
                batchCount += 1;
                const batchStartedAt = Date.now();

                // ── TEMPORARY diagnostic logging ─────────────────────
                // Logged synchronously for every still-active entity
                // BEFORE the Promise.all below is even constructed, so
                // all of a batch's START lines print together as one
                // group regardless of how long each entity's HTTP
                // response actually takes.
                if (customersTokens.length) console.log(`[BATCH ${batchCount}][Customers] START position=${startPosition} limit=${BATCH_SIZE}`);
                if (vendorsTokens.length) console.log(`[BATCH ${batchCount}][Vendors] START position=${startPosition} limit=${BATCH_SIZE}`);
                if (accountsTokens.length) console.log(`[BATCH ${batchCount}][Accounts] START position=${startPosition} limit=${BATCH_SIZE}`);
                if (classesTokens.length) console.log(`[BATCH ${batchCount}][Classes] START position=${startPosition} limit=${BATCH_SIZE}`);
                if (locationsTokens.length) console.log(`[BATCH ${batchCount}][Locations] START position=${startPosition} limit=${BATCH_SIZE}`);

                // All 5 entity APIs for this batch run together — nothing
                // here waits for another to finish first.
                const [customersResult, vendorsResult, accountsResult, classesResult, locationsResult] = await Promise.all([
                    customersTokens.length
                        ? QuickBooksService.getCustomersPage(customersTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                            .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(customersTokens) }))
                        : Promise.resolve(emptyPage()),
                    vendorsTokens.length
                        ? QuickBooksService.getVendorsPage(vendorsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                            .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(vendorsTokens) }))
                        : Promise.resolve(emptyPage()),
                    accountsTokens.length
                        ? QuickBooksService.getAccountsPage(accountsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                            .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(accountsTokens) }))
                        : Promise.resolve(emptyPage()),
                    classesTokens.length
                        ? QuickBooksService.getClassesPage(classesTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                            .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(classesTokens) }))
                        : Promise.resolve(emptyPage()),
                    locationsTokens.length
                        ? QuickBooksService.getLocationsPage(locationsTokens, startPosition, BATCH_SIZE, orgNameByTokenId)
                            .catch(() => ({ records: [], exhaustedTokenIds: allTokenIds(locationsTokens) }))
                        : Promise.resolve(emptyPage())
                ]);

                console.log(`[BATCH ${batchCount}][Customers] RESPONSE count=${customersResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
                console.log(`[BATCH ${batchCount}][Vendors] RESPONSE count=${vendorsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
                console.log(`[BATCH ${batchCount}][Accounts] RESPONSE count=${accountsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
                console.log(`[BATCH ${batchCount}][Classes] RESPONSE count=${classesResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);
                console.log(`[BATCH ${batchCount}][Locations] RESPONSE count=${locationsResult.records.length} (+${Date.now() - batchStartedAt}ms since this batch's requests started)`);

                customers.push(...customersResult.records);
                vendors.push(...vendorsResult.records);
                accounts.push(...accountsResult.records);
                classes.push(...classesResult.records);
                locations.push(...locationsResult.records);

                customersTokens = dropExhausted(customersTokens, customersResult.exhaustedTokenIds);
                vendorsTokens   = dropExhausted(vendorsTokens, vendorsResult.exhaustedTokenIds);
                accountsTokens  = dropExhausted(accountsTokens, accountsResult.exhaustedTokenIds);
                classesTokens   = dropExhausted(classesTokens, classesResult.exhaustedTokenIds);
                locationsTokens = dropExhausted(locationsTokens, locationsResult.exhaustedTokenIds);

                startPosition += BATCH_SIZE;
            }

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
     * Clears the authenticated user's own stored QuickBooks tokens only.
     */
    disconnectQuickbooks = async (req, res, next) => {
        try {
            await QuickBooksTokenRepository.clearTokens(req.user.email);
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
            const mail = req.user.email;
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
            const mail = req.user.email;
            const plan = req.query.plan || 'pro';

            const stats = {
                plan: plan.toLowerCase(),
                maxPerPlatform: 10,
                quickbooks: { connected: 0, remaining: 10 }
            };

            if (plan === 'trial')    stats.maxPerPlatform = 1;
            else if (plan === 'basic')    stats.maxPerPlatform = 1;
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
            const success = await QuickBooksService.disconnectConnection(companyId, req.user.email);
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
            const success = await QuickBooksService.activateConnection(companyId, req.user.email);
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

            const success = await QuickBooksService.renameConnection(companyId, req.user.email, companyName);
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

            const aggregated = await QuickBooksService.pullMasterData(companyId, tier, req.user.email);

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
                locations: aggregated.locations,
                isFirstSync: aggregated.isFirstSync
            });
        } catch (err) {
            return next(err);
        }
    };
}

module.exports = new QuickbooksController();
