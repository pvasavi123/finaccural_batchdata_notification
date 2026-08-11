'use strict';

const express    = require('express');
const router     = express.Router();
const controller = require('./controller');
const { authenticate } = require('../auth/auth.middleware');
const { validate } = require('../../core/middleware/validate');
const schemas = require('../../core/validation/schemas');

/**
 * Excel Validation Routes
 * -----------------------------------------------------------------
 * All routes are authenticated — comparisons against live API data or
 * DB rows are always scoped to req.user.email, never a client-supplied
 * identifier (same pattern as modules/quickbooks + modules/xero).
 * -----------------------------------------------------------------
 */

// POST /api/excel-validation/schema-check
router.post('/schema-check', authenticate, validate(schemas.excelSchemaCheck), controller.validateSchema);

// POST /api/excel-validation/data-type-check
router.post('/data-type-check', authenticate, validate(schemas.excelSchemaCheck), controller.validateDataTypes);

// POST /api/excel-validation/report[?format=xlsx]
router.post('/report', authenticate, validate(schemas.excelSchemaCheck), controller.generateReport);

// POST /api/excel-validation/vs-api
router.post('/vs-api', authenticate, validate(schemas.excelVsApi), controller.compareWithApi);

// POST /api/excel-validation/vs-database
router.post('/vs-database', authenticate, validate(schemas.excelVsDatabase), controller.compareWithDatabase);

module.exports = router;
