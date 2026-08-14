const path = require('path');
const dotenv = require('dotenv');

// Attempt loading .env from current directory, package directory, and explicitly configured paths
dotenv.config({ path: path.resolve(__dirname, '../../../.env') });
dotenv.config({ path: path.resolve(__dirname, '../../.env') });
dotenv.config({ path: path.resolve(__dirname, '../.env') });
dotenv.config({ path: path.resolve(process.cwd(), '.env') });
dotenv.config();

module.exports = {
    PORT: process.env.PORT || 8000,
    INSTANCE_ID: process.env.INSTANCE_ID || 'server-1',
    SESSION_SECRET: 'finaccural-node-secret',
    JWT_SECRET: process.env.JWT_SECRET || 'jwt-finaccural-secret-key-123',
    REDIS: {
        HOST: process.env.REDIS_HOST || 'localhost',
        PORT: parseInt(process.env.REDIS_PORT) || 6379,
        PASSWORD: process.env.REDIS_PASSWORD || undefined
    },
    DB: {
        HOST: process.env.DB_HOST || 'localhost',
        PORT: process.env.DB_PORT || 3306,
        NAME: process.env.DB_NAME || 'quickbooks_xero',
        USER: process.env.DB_USER || 'root',
        PASSWORD: process.env.DB_PASSWORD || '123456'
    },
    TRIAL: {
        // How long a brand-new signup's trial lasts, in milliseconds.
        // Defaults to 2 minutes for local testing; set TRIAL_DURATION_MS
        // in .env (e.g. 1209600000 for 14 days) for anything real.
        DURATION_MS:  parseInt(process.env.TRIAL_DURATION_MS, 10) || 2 * 60 * 1000,
        DEFAULT_PLAN: process.env.TRIAL_DEFAULT_PLAN || 'trial'
    },
    QB: {
        CLIENT_ID: process.env.QB_CLIENT_ID,
        CLIENT_SECRET: process.env.QB_CLIENT_SECRET,
        REDIRECT_URI: process.env.QB_REDIRECT_URI
    },
    XERO: {
        CLIENT_ID: process.env.XERO_CLIENT_ID,
        CLIENT_SECRET: process.env.XERO_CLIENT_SECRET,
        REDIRECT_URI: process.env.XERO_REDIRECT_URI,
        SCOPES: process.env.XERO_SCOPES || 'openid profile email offline_access accounting.contacts accounting.settings.read'
    },
    GOOGLE: {
        CLIENT_ID: process.env.GOOGLE_CLIENT_ID,
        CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET,
        REDIRECT_URI: process.env.GOOGLE_REDIRECT_URI || 'http://localhost:8000/api/auth/google/callback'
    },
    MICROSOFT: {
        CLIENT_ID: process.env.MICROSOFT_CLIENT_ID,
        CLIENT_SECRET: process.env.MICROSOFT_CLIENT_SECRET,
        REDIRECT_URI: process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:8000/api/microsoft/callback',
        // Azure AD tenant scope: a GUID for a single tenant, or 'common' / 'organizations' / 'consumers'.
        TENANT_ID: process.env.MICROSOFT_TENANT_ID || 'consumers',
        SCOPES: process.env.MICROSOFT_SCOPES || 'openid profile email User.Read offline_access'
    }
};
