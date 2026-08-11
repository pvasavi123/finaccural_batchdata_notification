'use strict';

/**
 * API Router
 * ------------------------------------------------------------------
 * This file mounts module routers and maps the unified global
 * connections and pull-master-data routes dynamically to the loaded
 * platform integrations (QuickBooks and Xero).
 *
 * This design is 100% loosely coupled: if QuickBooks or Xero is
 * commented out or removed for regional deployment, the endpoints
 * continue to function seamlessly without throwing module errors.
 * ------------------------------------------------------------------
 */

const express = require('express');
const router  = express.Router();
const { AppError } = require('../core/errors/AppError');
const { validate } = require('../core/middleware/validate');
const schemas = require('../core/validation/schemas');

// ── Core domain modules ───────────────────────────────────────────
const authRoutes        = require('../modules/auth/auth.routes');
const billingRoutes     = require('../modules/billing/billing.routes');

// ── Accounting integrations (conditionally imported / mounted) ────
let quickbooksRoutes;
try {
    quickbooksRoutes = require('../modules/quickbooks/routes');
} catch (e) {
    quickbooksRoutes = null;
}

let xeroRoutes;
try {
    xeroRoutes = require('../modules/xero/routes');
} catch (e) {
    xeroRoutes = null;
}

// ── Legacy / compatibility ────────────────────────────────────────
const adminRoutes = require('../modules/admin/routes');

// ── Mount Canonical Routers ───────────────────────────────────────

router.use('/auth', authRoutes);
router.use('/', billingRoutes);

if (quickbooksRoutes) {
    router.use('/quickbooks', quickbooksRoutes);
}
if (xeroRoutes) {
    router.use('/xero', xeroRoutes);
}

router.use('/admin', adminRoutes);

const authController = require('../modules/auth/auth.controller');

// Backward-compat aliases for OAuth providers
router.get('/google/connect',    authController.googleConnect);
router.get('/google/callback',   authController.googleCallback);
router.get('/microsoft/connect', authController.microsoftConnect);
router.get('/microsoft/callback', authController.microsoftCallback);

const { authenticate } = require('../modules/auth/auth.middleware');

// ── Dynamic Connections Routes ────────────────────────────────────
// ── Dynamic Connections Routes ────────────────────────────────────

// GET /api/connections
// The owning email is taken exclusively from the verified JWT
// (req.user.email, set by `authenticate`) — never from req.query.mail.
// Trusting a client-suppliable query param here would let any
// authenticated user list (and, via the endpoints below, disconnect/
// activate/rename/pull data for) another user's ERP connections just by
// passing their email.
router.get('/connections', authenticate, async (req, res, next) => {
    try {
        const mail = req.user.email;

        const list = [];
        
        if (quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            const qbList = await QuickBooksService.listConnections(mail);
            list.push(...qbList);
        }

        if (xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            const xeroList = await XeroService.listConnections(mail);
            list.push(...xeroList);
        }

        return res.json(list);
    } catch (err) {
        return next(err);
    }
});

// GET /api/connections/stats
router.get('/connections/stats', authenticate, validate(schemas.connectionStatsQuery, 'query'), async (req, res, next) => {
    try {
        const mail = req.user.email;
        const plan = req.query.plan || 'pro';

        const stats = {
            plan: plan.toLowerCase(),
            maxPerPlatform: 10,
            quickbooks: { connected: 0, remaining: 10 },
            xero:       { connected: 0, remaining: 10 }
        };

        if (plan === 'basic')    stats.maxPerPlatform = 1;
        else if (plan === 'standard') stats.maxPerPlatform = 3;

        if (quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            const qbStats = await QuickBooksService.getConnectionStats(mail, plan);
            stats.quickbooks = {
                connected: qbStats.connected,
                remaining: qbStats.remaining
            };
        }

        if (xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            const xeroStats = await XeroService.getConnectionStats(mail, plan);
            stats.xero = {
                connected: xeroStats.connected,
                remaining: xeroStats.remaining
            };
        }

        return res.json(stats);
    } catch (err) {
        return next(err);
    }
});

// DELETE /api/connections/:id
// `mail` (req.user.email) is passed through to the service layer so the
// update can only ever match a row this user actually owns — otherwise
// any authenticated user could disconnect another user's company just by
// knowing/guessing its companyId.
router.delete('/connections/:id', authenticate, async (req, res, next) => {
    try {
        const companyId = req.params.id;
        const mail = req.user.email;
        let success = false;

        if (quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            const qbSuccess = await QuickBooksService.disconnectConnection(companyId, mail);
            if (qbSuccess) success = true;
        }

        if (!success && xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            const xeroSuccess = await XeroService.disconnectConnection(companyId, mail);
            if (xeroSuccess) success = true;
        }

        return res.json({ success });
    } catch (err) {
        return next(err);
    }
});

// POST /api/connections/:id/activate
// Same ownership scoping as DELETE above.
router.post('/connections/:id/activate', authenticate, async (req, res, next) => {
    try {
        const companyId = req.params.id;
        const mail = req.user.email;
        let success = false;

        if (quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            const qbSuccess = await QuickBooksService.activateConnection(companyId, mail);
            if (qbSuccess) success = true;
        }

        if (!success && xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            const xeroSuccess = await XeroService.activateConnection(companyId, mail);
            if (xeroSuccess) success = true;
        }

        return res.json({ success });
    } catch (err) {
        return next(err);
    }
});

// PATCH /api/connections/:id/rename
// Same ownership scoping as DELETE above.
router.patch('/connections/:id/rename', authenticate, validate(schemas.renameConnection), async (req, res, next) => {
    try {
        const companyId = req.params.id;
        const mail = req.user.email;
        const { companyName } = req.body;

        let success = false;

        if (quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            const qbSuccess = await QuickBooksService.renameConnection(companyId, mail, companyName);
            if (qbSuccess) success = true;
        }

        if (!success && xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            const xeroSuccess = await XeroService.renameConnection(companyId, mail, companyName);
            if (xeroSuccess) success = true;
        }

        return res.json({ success });
    } catch (err) {
        return next(err);
    }
});

// GET /api/pull-master-data?companyId=...&platform=...&tier=...
// Same ownership scoping — a companyId this user doesn't own returns "not
// found" rather than silently pulling and returning another user's data.
router.get('/pull-master-data', authenticate, validate(schemas.pullMasterDataQuery, 'query'), async (req, res, next) => {
    try {
        const { companyId, platform, tier } = req.query;

        const mail = req.user.email;
        const normPlatform = platform.toLowerCase();
        let aggregated = null;

        if (normPlatform === 'quickbooks' && quickbooksRoutes) {
            const QuickBooksService = require('../modules/quickbooks/service');
            aggregated = await QuickBooksService.pullMasterData(companyId, tier, mail);
        } else if (normPlatform === 'xero' && xeroRoutes) {
            const XeroService = require('../modules/xero/service');
            aggregated = await XeroService.pullMasterData(companyId, tier, mail);
        }

        if (!aggregated) {
            throw new AppError('The requested resource was not found.', 404, 'ERR_NOT_FOUND', `No active connections found for ${platform}.`);
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
});

module.exports = router;
