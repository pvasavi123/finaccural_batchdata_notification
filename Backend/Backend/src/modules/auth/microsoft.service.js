'use strict';

const axios       = require('axios');
const querystring = require('querystring');
const logger      = require('../../core/logger');
const config      = require('../../core/config');

/**
 * MicrosoftAuthService
 * ----------------------------------------------------------------
 * Handles the low-level Microsoft Entra ID (Azure AD) OAuth 2.0 /
 * OpenID Connect HTTP calls against the Microsoft identity platform
 * v2.0 endpoint, plus the Microsoft Graph profile lookup.
 *
 * Mirrors google.service.js so AuthService/AuthController can treat
 * every OAuth provider the same way.
 * ----------------------------------------------------------------
 */
class MicrosoftAuthService {

    /**
     * Build the Microsoft Entra ID OAuth 2.0 authorisation redirect URL.
     *
     * @param {string} [loginHint] - A previously-seen account's email.
     *   When present, Microsoft is told which account to use via
     *   `login_hint`, and the forced `prompt: 'select_account'` is
     *   dropped — that's what lets the account-picker's "previous
     *   account" row skip straight past the account-chooser screen
     *   instead of asking the user to pick again. Without a hint (fresh
     *   "Add account") the chooser is still forced, same as before.
     * @returns {string}
     */
    getAuthUrl(loginHint) {
        const clientId    = config.MICROSOFT.CLIENT_ID || process.env.MICROSOFT_CLIENT_ID;
        const redirectUri = config.MICROSOFT.REDIRECT_URI || process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:8000/api/microsoft/callback';
        const tenantId    = config.MICROSOFT.TENANT_ID || process.env.MICROSOFT_TENANT_ID || 'consumers';
        const scopes      = config.MICROSOFT.SCOPES || 'openid profile email User.Read offline_access';

        if (!clientId) {
            logger.error('MICROSOFT_CLIENT_ID is missing from .env file!');
            throw new Error('MICROSOFT_CLIENT_ID is not configured in .env file');
        }

        const params = {
            client_id:     clientId,
            redirect_uri:  redirectUri,
            response_type: 'code',
            response_mode: 'query',
            scope:         scopes
        };

        if (loginHint) {
            params.login_hint = loginHint;
        } else {
            params.prompt = 'select_account';
        }

        return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${querystring.stringify(params)}`;
    }

    /**
     * Exchange an authorisation code for access/refresh/id tokens.
     * @param {string} code
     * @returns {Promise<object>} token response data
     */
    async exchangeCodeForToken(code) {
        try {
            const clientId     = config.MICROSOFT.CLIENT_ID || process.env.MICROSOFT_CLIENT_ID;
            const clientSecret = config.MICROSOFT.CLIENT_SECRET || process.env.MICROSOFT_CLIENT_SECRET;
            const redirectUri  = config.MICROSOFT.REDIRECT_URI || process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:8000/api/microsoft/callback';
            const tenantId     = config.MICROSOFT.TENANT_ID || process.env.MICROSOFT_TENANT_ID || 'consumers';
            const scopes       = config.MICROSOFT.SCOPES || 'openid profile email User.Read offline_access';

            const response = await axios.post(
                `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
                querystring.stringify({
                    code,
                    client_id:     clientId,
                    client_secret: clientSecret,
                    redirect_uri:  redirectUri,
                    grant_type:    'authorization_code',
                    scope:         scopes
                }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );
            return response.data;
        } catch (error) {
            logger.error('Failed to exchange Microsoft code for token', error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Fetch the authenticated user's Microsoft Graph profile.
     * @param {string} accessToken
     * @returns {Promise<{ sub, email, name }>} normalised to the same
     *          shape GoogleAuthService.getUserProfile() returns.
     */
    async getUserProfile(accessToken) {
        try {
            const response = await axios.get(
                'https://graph.microsoft.com/v1.0/me',
                { headers: { Authorization: `Bearer ${accessToken}` } }
            );

            const profile = response.data;
            // Personal Microsoft accounts (outlook.com/hotmail.com) may not
            // return `mail`; fall back to userPrincipalName as Google's
            // profile.sub / profile.email shape does for consistency.
            return {
                sub:   profile.id,
                email: profile.mail || profile.userPrincipalName,
                name:  profile.displayName || profile.mail || profile.userPrincipalName
            };
        } catch (error) {
            logger.error('Failed to fetch Microsoft user profile', error.response?.data || error.message);
            throw error;
        }
    }
}

module.exports = new MicrosoftAuthService();
