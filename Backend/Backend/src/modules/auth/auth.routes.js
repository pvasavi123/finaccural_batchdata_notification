'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./auth.controller');
const { authenticate } = require('./auth.middleware');
const { validate } = require('../../core/middleware/validate');
const schemas = require('../../core/validation/schemas');
const { authLimiter } = require('../../core/middleware/rateLimiters');

/**
 * Auth Routes
 * ----------------------------------------------------------------
 * All endpoints are mounted at /api/auth by routes/index.js
 *
 * Public:
 *   POST /api/auth/signup
 *   POST /api/auth/login
 *   POST /api/auth/refresh
 *   GET  /api/auth/google/connect
 *   GET  /api/auth/google/callback
 *   GET  /api/auth/microsoft/connect
 *   GET  /api/auth/microsoft/callback
 *   POST /api/auth/logout
 *
 * Protected (JWT required):
 *   GET  /api/auth/me
 *   POST /api/auth/update-plan
 *   POST /api/auth/start-trial
 * ----------------------------------------------------------------
 */

// Local auth
// authLimiter throttles brute-force login/signup attempts (Security
// Validation); validate() enforces the Joi schema (Data Type Validation)
// before the request ever reaches the controller.
router.post('/signup', authLimiter, validate(schemas.signup), controller.signup);
router.post('/login',  authLimiter, validate(schemas.login),  controller.login);

// Google OAuth
router.get('/google/connect',  controller.googleConnect);
router.get('/google/callback', controller.googleCallback);

// Microsoft Entra ID (Azure AD) OAuth
router.get('/microsoft/connect',  controller.microsoftConnect);
router.get('/microsoft/callback', controller.microsoftCallback);

// Token refresh — public; the whole point is to exchange an opaque
// refresh token for a new access token when the 15-minute JWT expires.
// Rate-limited to prevent brute-forcing the DB for valid refresh tokens.
router.post('/refresh', authLimiter, controller.refresh);

// Session teardown — intentionally NOT behind `authenticate`. The whole
// point is to let a client with an already-expired/invalid token still
// clear its server-side session; requiring a valid JWT here would 401
// exactly when a client most needs to call it.
router.post('/logout', controller.logout);

// Protected endpoints
router.get('/me', authenticate, controller.getMe);
router.post('/update-plan', authenticate, validate(schemas.updatePlan), controller.updatePlan);
// No body to validate — the trial is a fixed plan/duration set server-side
// from config.TRIAL, not something the client can parameterize.
router.post('/start-trial', authenticate, controller.startTrial);

module.exports = router;
