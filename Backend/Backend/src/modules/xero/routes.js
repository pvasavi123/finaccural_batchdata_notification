'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./controller');
const { validateXeroState } = require('../../core/middleware/oauthMiddleware');
const { authenticate } = require('../auth/auth.middleware');
const { validate } = require('../../core/middleware/validate');
const { oauthLimiter } = require('../../core/middleware/rateLimiters');
const schemas = require('../../core/validation/schemas');

// OAuth
// /connect requires auth so the connection is tagged with the verified
// req.user.email, never a client-suppliable ?mail= — the frontend already
// appends ?token=<jwt> to this URL for exactly this reason.
// oauthLimiter throttles repeated connect/callback attempts (Security
// Validation) — OAuth round trips are a common brute-force / abuse target.
router.get('/connect',             authenticate, oauthLimiter, validate(schemas.erpConnectQuery, 'query'), controller.connectXero);
router.get('/callback',            oauthLimiter, validateXeroState, controller.xeroCallback);
router.post('/select-companies',   validate(schemas.selectXeroCompanies), controller.selectCompanies);
router.post('/disconnect',         authenticate, controller.disconnectXero);
router.get(['/tokens', '/tokens/'], authenticate, controller.listXeroTokens);

// Data endpoints
router.get(['/contacts', '/contacts/'],         authenticate, controller.getContacts);
router.get(['/accounts', '/accounts/'],         authenticate, controller.getAccounts);
router.get(['/classes', '/classes/'],           authenticate, controller.getClasses);
router.get(['/locations', '/locations/'],       authenticate, controller.getLocations);
router.get(['/organisation', '/organisation/'], authenticate, controller.getOrganisation);

// Connection endpoints
router.get('/connections', authenticate, controller.listConnections);
router.get('/connections/stats', authenticate, validate(schemas.connectionStatsQuery, 'query'), controller.getConnectionStats);
router.delete('/connections/:id', authenticate, controller.disconnectConnection);
router.post('/connections/:id/activate', authenticate, controller.activateConnection);
router.patch('/connections/:id/rename', authenticate, validate(schemas.renameConnection), controller.renameConnection);
router.get('/pull-master-data', authenticate, validate(schemas.moduleMasterDataQuery, 'query'), controller.pullMasterData);

module.exports = router;