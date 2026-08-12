const express = require('express');
const cors = require('cors');
const helmet = require('helmet');


const routes = require('./routes');
const config = require('./core/config');
const { errorHandler, notFoundHandler } = require('./core/middleware/errorHandler');
const { validateContentType } = require('./core/middleware/validateHeaders');
const { sanitizeInput } = require('./core/middleware/sanitize');
const { responseTime } = require('./core/middleware/responseTime');
const { generalLimiter } = require('./core/middleware/rateLimiters');


const path = require("path");


const app = express();


// Performance Validation — times the FULL request lifecycle, so mount
// this before everything else.
app.use(responseTime);


// Security Validation — sets standard security-related response headers
// (X-Content-Type-Options, X-Frame-Options, Strict-Transport-Security,
// etc.). FinAccrual previously shipped Express's bare defaults with none
// of these set.
//
// Two of helmet's defaults are deliberately turned OFF because they
// would break real, existing functionality rather than harden it:
//   - contentSecurityPolicy: the Google/Microsoft OAuth popup pages
//     (modules/auth/views/oauthPopup.view.js) are raw HTML responses
//     with inline <script> blocks and a script tag loaded from
//     appsforoffice.microsoft.com — helmet's default CSP is
//     `script-src 'self'`, which would silently block all of it.
//   - crossOriginOpenerPolicy: those same popups talk back to the
//     Excel task pane via `window.opener.postMessage(...)`. Helmet's
//     default COOP (`same-origin`) can sever `window.opener` in the
//     popup, breaking that hand-off entirely.
app.use(helmet({
    contentSecurityPolicy: false,
    crossOriginOpenerPolicy: false
}));


app.use(cors());


// Header Validation — reject a JSON POST/PUT/PATCH whose Content-Type
// isn't application/json BEFORE express.json() silently no-ops on it.
app.use(validateContentType);


// Default express.json() limit is 100kb — too small for
// modules/excelValidation, which accepts an uploaded .xlsx workbook as
// base64 text inside the JSON body (see excelValidation/controller.js
// for why it's base64-in-JSON rather than multipart/form-data). Raised
// to 15mb everywhere else in this API a JSON body stays tiny (a few KB
// at most), so this only changes what excel-validation's routes can
// accept, not any other route's behavior.
app.use(express.json({ limit: '15mb' }));


// Security Validation — trims/strips control characters and inline
// script/event-handler HTML out of every string in req.body (and
// req.params); see core/middleware/sanitize.js for why req.query is
// intentionally excluded (Express 5 makes it non-reassignable).
app.use(sanitizeInput);


const session = require('express-session');
const { RedisStore } = require("connect-redis");
const redisClient = require('./core/redis');


app.use(session({
    store: new RedisStore({ client: redisClient }),
    secret: config.SESSION_SECRET || 'finaccrual-fallback-secret-key',
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: process.env.NODE_ENV === 'production',
        maxAge: 1000 * 60 * 60 * 24 // 24 hours
    }
}));


app.get('/favicon.ico', (req, res) => res.status(204).end());


// Security Validation — general request-rate ceiling across the whole
// API (stricter limiters for /api/auth/* and the QuickBooks/Xero OAuth
// routes are applied locally in their own route files).
app.use('/api', generalLimiter, routes);


app.use(notFoundHandler);
app.use(errorHandler);


module.exports = app;