'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./billing.controller');
const { authenticate } = require('../auth/auth.middleware');
const { validate } = require('../../core/middleware/validate');
const schemas = require('../../core/validation/schemas');

/**
 * Billing Routes
 * ------------------------------------------------------------------
 * Merges the old /api/subscription and /api/payments endpoints.
 *
 * Subscription:
 *   POST /api/subscription/upgrade  — JWT required
 *
 * Payments:
 *   GET  /api/payments/checkout     — public (popup, no JWT)
 *   POST /api/payments/complete     — public (called from checkout popup)
 * ------------------------------------------------------------------
 */

// ── Subscription sub-path ────────────────────────────────────────
router.post('/subscription/upgrade', authenticate, validate(schemas.billingUpgrade), controller.upgrade);

// ── Payments sub-path ────────────────────────────────────────────
router.get('/payments/checkout', authenticate, controller.checkout);
router.post('/payments/complete', authenticate, validate(schemas.completePayment), controller.completePayment);

module.exports = router;
