'use strict';

const axios        = require('axios');
const querystring  = require('querystring');
const config       = require('../../core/config');
const CONSTANTS    = require('../../core/constants');
const { encodeBasicAuth } = require('../../core/helpers');
const QuickBooksTokenRepository = require('./repository');
const QuickBooksMapper = require('./mapper');
const logger       = require('../../core/logger');
const QuickBooksTokenManager = require('./oauth/QuickBooksTokenManager');
const { ErpSessionExpiredError } = require('../../core/errors/AppError');

/**
 * QuickBooksService
 * -----------------------------------------------------------------
 * Responsible for all QuickBooks business logic:
 *   - OAuth token exchange & storage
 *   - Querying the QB API
 *   - Delegating data transformation to QuickBooksMapper
 * -----------------------------------------------------------------
 */
class QuickBooksService {

    /**
     * Exchange the OAuth authorization code for tokens, query CompanyInfo, and persist connection.
     * @param {string} code   - OAuth authorization code
     * @param {string} realmId - QB company ID
     */
    static async exchangeAndSaveToken(code, realmId, sessionInfo, mail) {
        const credentials = encodeBasicAuth(config.QB.CLIENT_ID, config.QB.CLIENT_SECRET);

        const response = await axios.post(
            CONSTANTS.QUICKBOOKS.TOKEN_URL,
            querystring.stringify({
                grant_type:   'authorization_code',
                code,
                redirect_uri: config.QB.REDIRECT_URI
            }),
            {
                headers: {
                    Accept:         'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded',
                    Authorization:  `Basic ${credentials}`
                }
            }
        );

        const tokenData = response.data;

        // Fetch CompanyInfo using the newly issued access token directly,
        // via a raw request rather than executeQuery/QuickBooksTokenManager.
        // Those look up whatever token is already stored for this realmId,
        // which — on a reconnect — can be a stale/expired one, causing this
        // step to fail with a 401 immediately after a successful exchange.
        let companyName = 'QuickBooks Company';
        try {
            const url = `${CONSTANTS.QUICKBOOKS.BASE_URL}/v3/company/${realmId}/query`;
            const compRes = await axios.get(url, {
                headers: {
                    Authorization: `Bearer ${tokenData.access_token}`,
                    Accept: 'application/json',
                    'Content-Type': 'application/text'
                },
                params: { query: 'SELECT * FROM CompanyInfo' }
            });
            const compInfo = QuickBooksMapper.toCompanyInfo(compRes.data);
            companyName = compInfo ? (compInfo.name || compInfo.legalName || realmId) : 'QuickBooks Company';
        } catch (compErr) {
            logger.warn(`Could not fetch company info directly during OAuth exchange for realm ${realmId}:`, compErr.message);
        }

        await QuickBooksTokenRepository.upsertToken({
            realm_id: realmId,
            access_token: tokenData.access_token || '',
            refresh_token: tokenData.refresh_token || '',
            token_type: tokenData.token_type || '',
            expires_in: Math.floor(Date.now() / 1000) + (tokenData.expires_in || 0),
            x_refresh_token_expires_in: Math.floor(Date.now() / 1000) + (tokenData.x_refresh_token_expires_in || 0),
            session_info: sessionInfo,
            mail: mail,
            company_name: companyName,
            // A freshly connected company hasn't had a Master Data Pull yet,
            // so it starts "Not Synced" rather than "Active" — pullMasterData
            // flips it to 'Active' once the first pull succeeds.
            status: 'Not Synced'
        });
    }

    /**
     * Execute a raw QBQL query against the QuickBooks API.
     * @param {string} query - QuickBooks SQL-like query string
     * @param {object} [token] - Specific QuickBooks Token record to use
     * @returns {object} raw API response
     */
    static async executeQuery(query, token) {
        let realmId;
        let accessToken;

        if (token) {
            realmId = token.companyId || token.realm_id;
            // Always go through QuickBooksTokenManager rather than falling
            // back to token.access_token/accessToken on failure. That token
            // is precisely the one getValidToken() just decided was
            // expiring/expired — silently using it anyway would mean a
            // failed/refused refresh (revoked connection) gets masked as a
            // doomed API call instead of surfacing as "Reconnect" here.
            accessToken = await QuickBooksTokenManager.getValidToken(realmId);
        } else {
            const connections = await QuickBooksTokenRepository.getActiveTokens();
            const activeToken = connections[0];
            if (!activeToken) {
                throw new ErpSessionExpiredError('QuickBooks', 'No active QuickBooks connection found.');
            }
            realmId = activeToken.companyId || activeToken.realm_id;
            accessToken = await QuickBooksTokenManager.getValidToken(realmId);
        }

        const url = `${CONSTANTS.QUICKBOOKS.BASE_URL}/v3/company/${realmId}/query`;

        try {
            const response = await axios.get(url, {
                headers: {
                    Authorization:  `Bearer ${accessToken}`,
                    Accept:         'application/json',
                    'Content-Type': 'application/text'
                },
                params: { query }
            });
            return response.data;
        } catch (error) {
            logger.error(`Error executing QB query for realm ${realmId}:`, error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Fetches every record of `entityName` for a token, paging through
     * STARTPOSITION/MAXRESULTS until QuickBooks returns fewer than
     * `batchSize` records.
     *
     * QuickBooks' query API defaults to only the first 100 records when
     * MAXRESULTS is omitted, and caps MAXRESULTS at 1000 per request, so
     * any entity type with more records than that gets silently truncated
     * unless paged through like this. Returns a QueryResponse-shaped
     * object so existing QuickBooksMapper.toXList() calls work unchanged.
     *
     * @param {string} entityName - QBQL entity name, e.g. "Customer".
     * @param {object} token
     * @param {number} [batchSize=1000] - QuickBooks' MAXRESULTS hard cap.
     * @returns {Promise<{ QueryResponse: Object }>}
     */
    static async queryAll(entityName, token, batchSize = 1000) {
        const fetchBatch = async (startPosition) => {
            const query = `SELECT * FROM ${entityName} STARTPOSITION ${startPosition} MAXRESULTS ${batchSize}`;
            const raw = await QuickBooksService.executeQuery(query, token);
            const batch = raw?.QueryResponse?.[entityName] || [];

            if (batch.length < batchSize) {
                return batch;
            }
            const nextBatch = await fetchBatch(startPosition + batchSize);
            return batch.concat(nextBatch);
        };

        const allRecords = await fetchBatch(1);
        return { QueryResponse: { [entityName]: allRecords } };
    }

    /**
     * Fetch company info and return clean CompanyDTO for a specific token or
     * all of the calling user's tokens.
     * @param {object} [token]
     * @param {string} [mail] - Owning user's email; scopes which companies are queried when `token` isn't given.
     * @returns {CompanyDTO|CompanyDTO[]|null}
     */
    static async getCompanyInfo(token, mail) {
        if (token) {
            try {
                const raw = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', token);
                return QuickBooksMapper.toCompanyInfo(raw);
            } catch (err) {
                return null;
            }
        }

        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        if (!tokens || tokens.length === 0) return null;

        const companyResults = await Promise.all(tokens.map(async (t) => {
            try {
                const raw = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', t);
                const info = QuickBooksMapper.toCompanyInfo(raw);
                if (info) {
                    info.id = t.companyId;
                    return info;
                }
            } catch (err) {}
            return null;
        }));

        const companies = companyResults.filter(Boolean);
        return companies.length === 1 ? companies[0] : companies;
    }

    /**
     * Helper to get company name and ID for tagging records of a specific token.
     * @param {object} token
     * @returns {Promise<{ orgId: string, orgName: string }>}
     */
    static async getCompanyMetadata(token) {
        const company = await QuickBooksService.getCompanyInfo(token).catch(() => null);
        const realmId = token.companyId || token.realm_id;
        const orgId   = company ? (company.id || company.name || realmId) : realmId;
        const orgName = company ? (company.name || company.legalName || company.id || "QuickBooks Company") : "QuickBooks Company";
        return { orgId, orgName };
    }

    /**
     * Fetch all customers and return clean CustomerDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {CustomerDTO[]}
     */
    static async getCustomers(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll('Customer', token);
                const list = QuickBooksMapper.toCustomerList(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(c => ({
                    ...c,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting customers for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all vendors and return clean VendorDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {VendorDTO[]}
     */
    static async getVendors(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll('Vendor', token);
                const list = QuickBooksMapper.toVendorList(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(v => ({
                    ...v,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting vendors for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all accounts and return clean AccountDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {AccountDTO[]}
     */
    static async getAccounts(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll('Account', token);
                const list = QuickBooksMapper.toAccountList(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(a => ({
                    ...a,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting accounts for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all classes and return clean ClassDTOs across the calling
     * user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {ClassDTO[]}
     */
    static async getClasses(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll('Class', token);
                const list = QuickBooksMapper.toClassList(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(c => ({
                    ...c,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting classes for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    /**
     * Fetch all locations (departments) and return clean LocationDTOs
     * across the calling user's connected companies only.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {LocationDTO[]}
     */
    static async getLocations(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const raw = await QuickBooksService.queryAll('Department', token);
                const list = QuickBooksMapper.toLocationList(raw);
                const { orgName } = await QuickBooksService.getCompanyMetadata(token);
                return list.map(l => ({
                    ...l,
                    clientId: orgName,
                    clientName: orgName
                }));
            } catch (err) {
                const realmId = token.companyId || token.realm_id;
                logger.error(`Error getting departments for realm ${realmId}:`, err.message);
                return [];
            }
        }));
        return results.flat();
    }

    // ── Self-contained Connections Management & Pulling ────────────────

    static PLAN_LIMITS = { trial: 1, basic: 1, standard: 3, pro: 10 };

    static getMaxConnections(plan) {
        return QuickBooksService.PLAN_LIMITS[(plan || 'pro').toLowerCase()] ?? 10;
    }

    static async listConnections(mail) {
        const { QuickBooksToken } = require('../../core/database');
        const qbWhere = mail ? { mail } : {};
        const qbTokens = await QuickBooksToken.findAll({ where: qbWhere });

        return qbTokens.map(t => ({
            platform:     'QuickBooks',
            companyName:  t.company_name || 'QuickBooks Company',
            companyId:    t.realm_id,
            status:       t.status || 'Not Synced',
            lastSyncedAt: t.last_synced_at || t.updated_at || null,
            createdAt:    t.created_at || null
        }));
    }

    static async getConnectionStats(mail, plan) {
        const { QuickBooksToken } = require('../../core/database');
        const { Op } = require('sequelize');
        const maxAllowed = QuickBooksService.getMaxConnections(plan);

        const whereClause = { status: { [Op.ne]: 'Disconnected' } };
        if (mail) whereClause.mail = mail;

        const qbCount = await QuickBooksToken.count({ where: whereClause });

        return {
            plan: (plan || 'pro').toLowerCase(),
            maxAllowed,
            connected: qbCount,
            remaining: Math.max(0, maxAllowed - qbCount)
        };
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required: without it this
     *   would disconnect a company regardless of who owns it, letting any
     *   authenticated user tear down another user's connection just by
     *   knowing/guessing its companyId.
     */
    static async disconnectConnection(companyId, mail) {
        if (!mail) return false;
        const { QuickBooksToken } = require('../../core/database');
        const [updated] = await QuickBooksToken.update(
            { status: 'Disconnected' },
            { where: { realm_id: companyId, mail } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async activateConnection(companyId, mail) {
        if (!mail) return false;
        const { QuickBooksToken } = require('../../core/database');
        const { Op } = require('sequelize');

        // Selecting/switching to a connection re-activates it if it was
        // 'Disconnected', but must not resurrect 'Not Synced' to 'Active' —
        // that transition only happens via a successful Master Data Pull
        // (see pullMasterData).
        const [updated] = await QuickBooksToken.update(
            { status: 'Active' },
            { where: { realm_id: companyId, mail, status: { [Op.ne]: 'Not Synced' } } }
        );
        if (updated > 0) return true;

        // If nothing matched, the row might legitimately be 'Not Synced'
        // (or simply not exist) — confirm it exists (and is owned by this
        // user) so the caller still gets a truthy result for "this company
        // is now the active one".
        const existing = await QuickBooksToken.findOne({ where: { realm_id: companyId, mail } });
        return !!existing;
    }

    /**
     * @param {string} companyId
     * @param {string} mail - Owning user's email. Required — see
     *   disconnectConnection() above for why an ownership check matters here.
     */
    static async renameConnection(companyId, mail, companyName) {
        if (!mail) return false;
        const { QuickBooksToken } = require('../../core/database');
        const [updated] = await QuickBooksToken.update(
            { company_name: companyName },
            { where: { realm_id: companyId, mail } }
        );
        return updated > 0;
    }

    /**
     * @param {string} companyId
     * @param {string} tier
     * @param {string} mail - Owning user's email. Required — without it a
     *   companyId-scoped pull would return (and let this user overwrite
     *   their Excel sheet with) another user's financial data, and a
     *   bulk (no companyId) pull would aggregate every user's connections
     *   in the system into one response.
     */
    static async pullMasterData(companyId, tier, mail) {
        if (!mail) return null;
        const { QuickBooksToken } = require('../../core/database');
        const { Op } = require('sequelize');
        const maxAllowed = QuickBooksService.getMaxConnections(tier);

        // Exclude 'Disconnected' connections from the bulk (no companyId)
        // pull — a revoked/expired connection stops being retried
        // automatically the moment it's marked disconnected; it only comes
        // back once the user reconnects. 'Not Synced' connections are still
        // included since they've never had a chance to sync yet. Both
        // branches are scoped to `mail` so this can only ever touch the
        // calling user's own companies.
        const rawTokens = companyId
            ? await QuickBooksToken.findAll({ where: { realm_id: companyId, mail } })
            : await QuickBooksToken.findAll({ where: { mail, status: { [Op.ne]: 'Disconnected' } }, order: [['updated_at', 'DESC']] });

        const tokens = rawTokens.slice(0, maxAllowed).map(t => ({
            platform:     'quickbooks',
            companyId:    t.realm_id,
            companyName:  t.company_name || 'QuickBooks Company',
            realm_id:     t.realm_id,
            lastSyncedAt: t.last_synced_at
        }));

        const results = await Promise.all(tokens.map(async (token) => {
            try {
                const rawComp = await QuickBooksService.executeQuery('SELECT * FROM CompanyInfo', token);
                const comp = QuickBooksMapper.toCompanyInfo(rawComp);
                const companyList = comp ? [{ ...comp, id: token.companyId }] : [];

                const [rawCust, rawVend, rawAcc, rawClass, rawLoc] = await Promise.all([
                    QuickBooksService.queryAll('Customer', token),
                    QuickBooksService.queryAll('Vendor', token),
                    QuickBooksService.queryAll('Account', token),
                    QuickBooksService.queryAll('Class', token),
                    QuickBooksService.queryAll('Department', token)
                ]);

                const orgName = comp?.name || comp?.legalName || token.companyName;
                const tag = (list) => list.map(i => ({ ...i, clientId: orgName, clientName: orgName }));

                // Captured before the update below overwrites it — this is
                // what lets the frontend's Refresh Schedule flow tell "the
                // very first sync for this connection" (nothing to append
                // against yet, write everything) apart from "a later
                // refresh where nothing happens to be new" (both look
                // identical if you only look at isNew flags, since every
                // record's isNew is false in both cases).
                const isFirstSync = !token.lastSyncedAt;

                await QuickBooksToken.update(
                    { last_synced_at: new Date(), status: 'Active' },
                    { where: { realm_id: token.companyId } }
                );

                return {
                    company: companyList,
                    customers: tag(QuickBooksMapper.toCustomerList(rawCust, token.lastSyncedAt)),
                    vendors: tag(QuickBooksMapper.toVendorList(rawVend, token.lastSyncedAt)),
                    accounts: tag(QuickBooksMapper.toAccountList(rawAcc, token.lastSyncedAt)),
                    classes: tag(QuickBooksMapper.toClassList(rawClass, token.lastSyncedAt)),
                    locations: tag(QuickBooksMapper.toLocationList(rawLoc, token.lastSyncedAt)),
                    isFirstSync
                };
            } catch (err) {
                logger.error(`Error pulling QB data for connection ${token.companyId}:`, err.message);

                const isTokenError = err.response?.status === 401
                    || err.statusCode === 401
                    || (err.message && (err.message.includes('Token expired') || err.message.includes('401') || err.message.includes('grant')));

                if (isTokenError || err.message?.includes('OAuth')) {
                    await QuickBooksToken.update(
                        { status: 'Disconnected' },
                        { where: { realm_id: token.companyId } }
                    );
                    throw new ErpSessionExpiredError(
                        'QuickBooks',
                        `QuickBooks refresh token expired/revoked for company "${token.companyName}" (${token.companyId}): ${err.message}`
                    );
                }

                throw err;
            }
        }));

        return results.reduce((acc, curr) => ({
            company: [...acc.company, ...curr.company],
            customers: [...acc.customers, ...curr.customers],
            vendors: [...acc.vendors, ...curr.vendors],
            accounts: [...acc.accounts, ...curr.accounts],
            classes: [...acc.classes, ...curr.classes],
            locations: [...acc.locations, ...curr.locations],
            // A companyId-scoped pull (the normal case) is always exactly
            // one connection, so this just reflects that one token. For the
            // rare bulk (no companyId) pull spanning several connections,
            // AND-merging means "first sync" only holds if every one of
            // them is — mixing "never synced" and "already synced" here
            // would otherwise leave it ambiguous which single answer to
            // give the frontend for a mixed batch.
            isFirstSync: acc.isFirstSync && curr.isFirstSync
        }), { company: [], customers: [], vendors: [], accounts: [], classes: [], locations: [], isFirstSync: true });
    }
}

// Register event listener for plan downgrades
const eventBus = require('../../core/events');
const { QuickBooksToken } = require('../../core/database');

eventBus.on('user.downgraded', async ({ email }) => {
    try {
        const deletedCount = await QuickBooksToken.destroy({ where: { mail: email } });
        logger.info(`[QuickBooksService] Plan downgrade: cleared ${deletedCount} connections for ${email}`);
    } catch (err) {
        logger.error(`[QuickBooksService] Failed to clear connections on downgrade for ${email}:`, err.message);
    }
});

module.exports = QuickBooksService;
