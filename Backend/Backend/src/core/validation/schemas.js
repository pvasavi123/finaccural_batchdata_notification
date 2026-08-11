'use strict';

const Joi = require('joi');
const { SHEET_SCHEMAS, API_COMPARABLE_SHEETS } = require('../../modules/excelValidation/schemas/masterDataSchemas');

/**
 * Joi Schemas
 * ----------------------------------------------------------------
 * Data Type Validation.
 *
 * Centralized request-shape schemas, used with core/middleware/validate.js.
 * These sit in front of the hand-rolled checks that already exist in
 * modules/auth/auth.validation.js (regex email / min-length password) —
 * the difference is that Joi also enforces actual data TYPES ("email
 * must be a string", "tier must be one of these three strings", "an
 * extra unexpected field is stripped/rejected") which the regex-only
 * checks never covered.
 * ----------------------------------------------------------------
 */

const email = Joi.string().trim().lowercase().email({ tlds: false }).max(150).required();

const signup = Joi.object({
    name:     Joi.string().trim().min(1).max(100).allow('', null),
    email,
    password: Joi.string().min(6).max(200).required()
});

const login = Joi.object({
    email,
    password: Joi.string().min(1).max(200).required()
});

const updatePlan = Joi.object({
    plan: Joi.string().trim().lowercase().valid('basic', 'standard', 'pro').required()
});

const renameConnection = Joi.object({
    companyName: Joi.string().trim().min(1).max(255).required()
});

// GET /api/pull-master-data?companyId=...&platform=...&tier=...
const pullMasterDataQuery = Joi.object({
    companyId: Joi.string().trim().max(255).allow('', null),
    platform:  Joi.string().trim().lowercase().valid('quickbooks', 'xero').required(),
    tier:      Joi.string().trim().lowercase().valid('basic', 'standard', 'pro').default('pro')
});

// GET /api/quickbooks/pull-master-data?companyId=...&tier=...
// GET /api/xero/pull-master-data?companyId=...&tier=...
// (module-scoped variant — no `platform` param, since the module is
// already implied by which router this is mounted under)
const moduleMasterDataQuery = Joi.object({
    companyId: Joi.string().trim().max(255).allow('', null),
    tier:      Joi.string().trim().lowercase().valid('basic', 'standard', 'pro').default('pro')
});

// GET /api/connections/stats?plan=...
const connectionStatsQuery = Joi.object({
    plan: Joi.string().trim().lowercase().valid('basic', 'standard', 'pro').default('pro')
});

// GET /api/quickbooks/connect?tier=... , GET /api/xero/connect?tier=...
const erpConnectQuery = Joi.object({
    tier: Joi.string().trim().lowercase().valid('basic', 'standard', 'pro').default('pro')
}).unknown(true); // OAuth connect URLs may legitimately carry other client-added params

// POST /api/admin/login
const adminLogin = Joi.object({
    email,
    password: Joi.string().min(1).max(200).required()
});

// POST /api/admin/signup
const adminSignup = Joi.object({
    name:     Joi.string().trim().min(1).max(100).required(),
    email,
    password: Joi.string().min(6).max(200).required()
});

// modules/billing/billing.service.js BillingService.ALLOWED_PLANS is
// EXACT-CASE ['Basic', 'Standard', 'Pro'] (isValidPlan does a
// case-sensitive `.includes()`) — deliberately NOT lowercased here like
// the other plan schemas above, or every billing request would fail
// isValidPlan() even though it passed this schema.
const billingPlan = Joi.string().trim().valid('Basic', 'Standard', 'Pro').required();

// POST /api/subscription/upgrade — { plan }
const billingUpgrade = Joi.object({
    plan: billingPlan
});

// POST /api/payments/complete — { email, plan }. Distinct from
// billingUpgrade because this one is called from the checkout popup and
// has to identify the user by email in the body instead of req.user.
const completePayment = Joi.object({
    email,
    plan: billingPlan
});

// POST /api/xero/select-companies — { selectedTenantIds: [...] }
const selectXeroCompanies = Joi.object({
    selectedTenantIds: Joi.array().items(Joi.string().trim().min(1).max(255)).min(1).required()
});

// The uploaded workbook travels as base64 inside the JSON body (see
// modules/excelValidation/controller.js for why — the multipart
// alternative would require carving an exception into the global
// validateContentType header gate). 15MB of base64 text comfortably
// covers a multi-thousand-row master-data workbook; app.js raises the
// express.json() body-size limit to match.
const fileBase64 = Joi.string().min(4).max(20 * 1024 * 1024).required();

// Every individual sheet schema this module knows about, plus the
// special "MasterData" key that validates the full multi-sheet
// workbook (Company/Customers/Vendors/Accounts/Classes/Locations) at
// once — see modules/excelValidation/service.js#resolveSchemas.
const excelSchemaNames = [...Object.keys(SHEET_SCHEMAS), 'MasterData'];

// POST /api/excel-validation/schema-check
// POST /api/excel-validation/data-type-check
// POST /api/excel-validation/report
const excelSchemaCheck = Joi.object({
    fileBase64,
    schema: Joi.string().valid(...excelSchemaNames).required()
});

// POST /api/excel-validation/vs-api — { fileBase64, sheet, platform }
const excelVsApi = Joi.object({
    fileBase64,
    sheet:    Joi.string().valid(...API_COMPARABLE_SHEETS).required(),
    platform: Joi.string().trim().lowercase().valid('quickbooks', 'xero').required()
});

// POST /api/excel-validation/vs-database — { fileBase64 }
const excelVsDatabase = Joi.object({
    fileBase64
});

module.exports = {
    signup,
    login,
    updatePlan,
    renameConnection,
    pullMasterDataQuery,
    moduleMasterDataQuery,
    connectionStatsQuery,
    erpConnectQuery,
    adminLogin,
    adminSignup,
    billingUpgrade,
    completePayment,
    selectXeroCompanies,
    excelSchemaCheck,
    excelVsApi,
    excelVsDatabase
};
