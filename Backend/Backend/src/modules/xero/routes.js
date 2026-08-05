'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./controller');
const { validateXeroState } = require('../../core/middleware/oauthMiddleware');
const { authenticate } = require('../auth/auth.middleware');

// OAuth
router.get('/connect',             controller.connectXero);
router.get('/callback',            validateXeroState, controller.xeroCallback);
router.post('/select-companies',   controller.selectCompanies);
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
router.get('/connections/stats', authenticate, controller.getConnectionStats);
router.delete('/connections/:id', authenticate, controller.disconnectConnection);
router.post('/connections/:id/activate', authenticate, controller.activateConnection);
router.patch('/connections/:id/rename', authenticate, controller.renameConnection);
router.get('/pull-master-data', authenticate, controller.pullMasterData);

module.exports = router;