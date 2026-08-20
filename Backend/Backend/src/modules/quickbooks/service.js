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
const { AppError, ErpSessionExpiredError } = require('../../core/errors/AppError');

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

        // ── TEMPORARY concurrency-verification instrumentation ──────────
        // Logs the actual token-resolution + HTTP round trip timing for
        // every QuickBooks query, tagged with entity/STARTPOSITION parsed
        // straight out of the QBQL string, so REQUEST START timestamps
        // from different calls (e.g. Customer @1, Vendor @1, Account @1
        // fired from the same Promise.all) can be compared directly to
        // prove — from real wall-clock timestamps, not code inspection —
        // whether they were genuinely concurrent or serialized somewhere
        // beneath the service layer (token lookup, DB pool, axios/agent,
        // etc). Safe to delete once concurrency is confirmed.
        const qbEntityMatch = /FROM\s+(\w+)/i.exec(query);
        const qbPosMatch = /STARTPOSITION\s+(\d+)/i.exec(query);
        const qbLabel = `${qbEntityMatch ? qbEntityMatch[1] : 'query'}${qbPosMatch ? ` @${qbPosMatch[1]}` : ''}`;
        const qbCallStart = Date.now();

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

        console.log(`[QB-HTTP] ${new Date().toISOString()} TOKEN READY     ${qbLabel} realm=${realmId} (+${Date.now() - qbCallStart}ms since executeQuery() was called)`);

        const url = `${CONSTANTS.QUICKBOOKS.BASE_URL}/v3/company/${realmId}/query`;

        const qbHttpStart = Date.now();
        console.log(`[QB-HTTP] ${new Date(qbHttpStart).toISOString()} REQUEST START   ${qbLabel} realm=${realmId}`);

        try {
            const response = await axios.get(url, {
                headers: {
                    Authorization:  `Bearer ${accessToken}`,
                    Accept:         'application/json',
                    'Content-Type': 'application/text'
                },
                params: { query }
            });
            console.log(`[QB-HTTP] ${new Date().toISOString()} RESPONSE OK     ${qbLabel} realm=${realmId} (+${Date.now() - qbHttpStart}ms)`);
            return response.data;
        } catch (error) {
            console.log(`[QB-HTTP] ${new Date().toISOString()} RESPONSE ERROR  ${qbLabel} realm=${realmId} (+${Date.now() - qbHttpStart}ms)`);
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
     * Fetches exactly ONE page of `entityName` for a single token via
     * STARTPOSITION/MAXRESULTS — the single-page counterpart to
     * queryAll() above, which recurses through every page internally
     * and returns everything at once. This is what lets a caller (see
     * exportMasterData's batch loop) drive pagination one 10-record
     * batch at a time instead of waiting for a full recursive fetch.
     * @param {string} entityName - QBQL entity name, e.g. "Customer".
     * @param {object} token
     * @param {number} startPosition - 1-based, QuickBooks STARTPOSITION.
     * @param {number} pageSize - QuickBooks MAXRESULTS for this page.
     * @returns {Promise<{ raw: object, records: object[], hasMore: boolean }>}
     */
    static async queryPage(entityName, token, startPosition, pageSize) {
        const query = `SELECT * FROM ${entityName} STARTPOSITION ${startPosition} MAXRESULTS ${pageSize}`;
        const raw = await QuickBooksService.executeQuery(query, token);
        const records = raw?.QueryResponse?.[entityName] || [];
        // A page shorter than pageSize means this was the entity's last
        // page for this token — same "fewer than batchSize" signal
        // queryAll() uses internally to stop recursing.
        return { raw, records, hasMore: records.length === pageSize };
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
     * Fetches every active token's CompanyInfo ONCE and returns both the
     * "Company" worksheet data and a token→orgName lookup in a single
     * pass. Used by exportMasterData's batch loop so its per-batch,
     * per-entity page fetches (getCustomersPage/getVendorsPage/etc.,
     * below) can reuse an already-known org name instead of each one
     * separately re-fetching CompanyInfo per token per batch — the way
     * getCustomers()/getVendors()/etc. do today (once per entity type,
     * via getCompanyMetadata) is fine for a single non-paginated call,
     * but would mean 5x redundant CompanyInfo calls per token per batch
     * if repeated on every page.
     * @param {string} mail - Owning user's email; scopes which companies are queried.
     * @returns {Promise<{ tokens: object[], company: CompanyDTO|CompanyDTO[]|null, orgNameByTokenId: Map<string, string> }>}
     */
    static async getCompanyInfoAndOrgNames(mail) {
        const tokens = await QuickBooksTokenRepository.getActiveTokens(mail);
        const orgNameByTokenId = new Map();

        // Match getCompanyInfo(undefined, mail)'s own no-connections
        // behavior exactly (returns null, not []) — exportMasterData's
        // `if (company) {...}` check depends on that, and an empty
        // array is truthy in JS.
        if (!tokens || tokens.length === 0) {
            return { tokens: [], company: null, orgNameByTokenId };
        }

        const companies = (await Promise.all(tokens.map(async (token) => {
            const tokenId = token.companyId || token.realm_id;
            const info = await QuickBooksService.getCompanyInfo(token).catch(() => null);
            if (info) info.id = token.companyId;
            const orgName = info ? (info.name || info.legalName || info.id || "QuickBooks Company") : "QuickBooks Company";
            orgNameByTokenId.set(tokenId, orgName);
            return info;
        }))).filter(Boolean);

        const company = companies.length === 1 ? companies[0] : companies;
        return { tokens, company, orgNameByTokenId };
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

    // ── Paginated (batch) entity fetchers ───────────────────────────────
    //
    // Counterparts to getCustomers/getVendors/getAccounts/getClasses/
    // getLocations above, for callers that want to drive pagination one
    // page at a time (see exportMasterData's batch loop) instead of
    // getting every record back in a single call. Each one fetches at
    // most `pageSize` records per still-active token, in parallel, and
    // reports back which tokens have run out of pages so the caller can
    // stop querying them on the next batch — no per-token pagination
    // state is kept here between calls, that's the caller's job.

    /**
     * Shared implementation behind getCustomersPage/getVendorsPage/etc.
     * @param {string} entityName - QBQL entity, e.g. "Customer".
     * @param {Function} mapperFn - QuickBooksMapper.toXList, e.g. toCustomerList.
     * @param {object[]} activeTokens - Tokens still known to have more pages for this entity.
     * @param {number} startPosition - 1-based, QuickBooks STARTPOSITION.
     * @param {number} pageSize - QuickBooks MAXRESULTS for this page.
     * @param {Map<string,string>} orgNameByTokenId - From getCompanyInfoAndOrgNames(), so each
     *   page doesn't need its own CompanyInfo round trip just to tag clientId/clientName.
     * @returns {Promise<{ records: object[], exhaustedTokenIds: Set<string> }>}
     */
    static async _queryEntityPage(entityName, mapperFn, activeTokens, startPosition, pageSize, orgNameByTokenId) {
        const exhaustedTokenIds = new Set();
        const perToken = await Promise.all(activeTokens.map(async (token) => {
            const tokenId = token.companyId || token.realm_id;
            try {
                const { raw, records, hasMore } = await QuickBooksService.queryPage(entityName, token, startPosition, pageSize);
                if (!hasMore) exhaustedTokenIds.add(tokenId);
                const list = mapperFn(raw);
                const orgName = (orgNameByTokenId && orgNameByTokenId.get(tokenId)) || "QuickBooks Company";
                return list.map(item => ({ ...item, clientId: orgName, clientName: orgName }));
            } catch (err) {
                logger.error(`Error getting ${entityName} page (start ${startPosition}) for realm ${tokenId}:`, err.message);
                // A token whose page request failed isn't retried on the
                // next batch — same "log it, return empty, move on"
                // contract as the non-paginated getters' per-token catch.
                exhaustedTokenIds.add(tokenId);
                return [];
            }
        }));
        return { records: perToken.flat(), exhaustedTokenIds };
    }

    /** Paginated counterpart to getCustomers() — see _queryEntityPage. */
    static async getCustomersPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Customer', QuickBooksMapper.toCustomerList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getVendors() — see _queryEntityPage. */
    static async getVendorsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Vendor', QuickBooksMapper.toVendorList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getAccounts() — see _queryEntityPage. */
    static async getAccountsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Account', QuickBooksMapper.toAccountList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getClasses() — see _queryEntityPage. */
    static async getClassesPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Class', QuickBooksMapper.toClassList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /** Paginated counterpart to getLocations() — see _queryEntityPage. */
    static async getLocationsPage(activeTokens, startPosition, pageSize, orgNameByTokenId) {
        return QuickBooksService._queryEntityPage('Department', QuickBooksMapper.toLocationList, activeTokens, startPosition, pageSize, orgNameByTokenId);
    }

    /**
     * Fetches every page of the 5 paginated entities (Customer, Vendor,
     * Account, Class, Department) for a SINGLE token, `pageSize` records
     * at a time — the same concurrent-per-batch shape as exportMasterData's
     * controller loop (all 5 requests for the current batch fire together
     * via Promise.all, and the next batch only starts once that one
     * resolves), just scoped to one token instead of driving multi-token
     * exhaustion tracking across a whole connection list.
     *
     * This is what pullMasterData (the endpoint the Excel Add-in's Pull
     * Master Data / Refresh Schedule buttons actually call) uses instead
     * of firing one Promise.all of 5 full recursive queryAll() calls —
     * queryAll's own internal paging (MAXRESULTS up to 1000) never
     * produced the batch-by-batch, 10-records-at-a-time concurrency the
     * QuickBooks batch spec calls for, even though the 5 entity types
     * were already running concurrently relative to each other.
     *
     * Each entity independently drops out of later batches once it
     * returns a short page — e.g. Vendor can stop after 15 records while
     * Customer, with 47, keeps paging — without affecting the others.
     * Returns QueryResponse-shaped objects so the existing
     * QuickBooksMapper.toXList(raw, lastSyncedAt) calls in pullMasterData
     * work completely unchanged (isNew/isUpdated flagging included).
     *
     * @param {object} token
     * @param {number} [pageSize=10]
     * @returns {Promise<{ Customer: object[], Vendor: object[], Account: object[], Class: object[], Department: object[] }>}
     */
    static async _fetchAllPaginatedEntitiesForToken(token, pageSize = 10) {
        const entities = ['Customer', 'Vendor', 'Account', 'Class', 'Department'];
        const recordsByEntity = { Customer: [], Vendor: [], Account: [], Class: [], Department: [] };
        let active = entities.slice();
        let startPosition = 1;

        while (active.length > 0) {
            // All still-active entities for this token start together —
            // nothing here waits for another to finish first.
            const pages = await Promise.all(active.map(entityName =>
                QuickBooksService.queryPage(entityName, token, startPosition, pageSize)
            ));

            const stillActive = [];
            active.forEach((entityName, i) => {
                const { records, hasMore } = pages[i];
                recordsByEntity[entityName].push(...records);
                if (hasMore) stillActive.push(entityName);
            });
            active = stillActive;

            startPosition += pageSize;
        }

        return recordsByEntity;
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

                // Concurrent, 10-records-per-batch fetch across all 5
                // paginated entities for this token — see
                // _fetchAllPaginatedEntitiesForToken. CompanyInfo above is
                // deliberately outside this loop: it's a single record per
                // company, not paginated.
                const pagedEntities = await QuickBooksService._fetchAllPaginatedEntitiesForToken(token, 10);
                const rawCust  = { QueryResponse: { Customer:   pagedEntities.Customer } };
                const rawVend  = { QueryResponse: { Vendor:     pagedEntities.Vendor } };
                const rawAcc   = { QueryResponse: { Account:    pagedEntities.Account } };
                const rawClass = { QueryResponse: { Class:      pagedEntities.Class } };
                const rawLoc   = { QueryResponse: { Department: pagedEntities.Department } };

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

                // Check for QuickBooks Subscription Expired / Suspended (Code 8020)
                const faultError = err.response?.data?.Fault?.Error?.[0];
                if (faultError?.code === '8020' || (faultError?.Message && faultError.Message.includes('Subscription is not active'))) {
                    await QuickBooksToken.update(
                        { status: 'Disconnected' },
                        { where: { realm_id: token.companyId } }
                    );
                    throw new AppError(
                        'Your QuickBooks subscription has expired or been suspended. Please log into QuickBooks to update your billing.',
                        403,
                        'ERR_QB_SUBSCRIPTION_EXPIRED'
                    );
                }

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
