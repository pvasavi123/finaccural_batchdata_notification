'use strict';

const axios       = require('axios');
const querystring = require('querystring');
const logger      = require('../../core/logger');
const config      = require('../../core/config');

/**
 * GoogleAuthService
 * ----------------------------------------------------------------
 * Handles the low-level Google OAuth 2.0 HTTP calls.
 * Moved from modules/google/service.js so that all auth-related
 * code lives inside the auth module.
 *
 * The original modules/google/service.js is kept intact and simply
 * re-exports this file to avoid breaking existing imports.
 * ----------------------------------------------------------------
 */
class GoogleAuthService {

    /**
     * Build the Google OAuth 2.0 authorisation redirect URL.
     * @returns {string}
     */
    getAuthUrl() {
        const clientId    = config.GOOGLE.CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
        const redirectUri = config.GOOGLE.REDIRECT_URI || process.env.GOOGLE_REDIRECT_URI || 'http://localhost:8000/api/auth/google/callback';

        if (!clientId) {
            logger.error('GOOGLE_CLIENT_ID is missing from .env file!');
            throw new Error('GOOGLE_CLIENT_ID is not configured in .env file');
        }

        const params = {
            client_id:     clientId,
            redirect_uri:  redirectUri,
            response_type: 'code',
            scope:         'openid email profile',
            access_type:   'offline',
            prompt:        'consent'
        };
        return `https://accounts.google.com/o/oauth2/v2/auth?${querystring.stringify(params)}`;
    }

    /**
     * Exchange an authorisation code for access/refresh tokens.
     * @param {string} code
     * @returns {Promise<object>} token response data
     */
    async exchangeCodeForToken(code) {
        try {
            const clientId     = config.GOOGLE.CLIENT_ID || process.env.GOOGLE_CLIENT_ID;
            const clientSecret = config.GOOGLE.CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET;
            const redirectUri  = config.GOOGLE.REDIRECT_URI || process.env.GOOGLE_REDIRECT_URI || 'http://localhost:8000/api/auth/google/callback';

            console.log("EXCHANGE PARAMS:", {
                code: code ? 'Present' : 'Missing',
                client_id: clientId,
                client_secret_len: clientSecret ? clientSecret.length : 0,
                redirect_uri: redirectUri
            });

            const response = await axios.post(
                'https://oauth2.googleapis.com/token',
                querystring.stringify({
                    code,
                    client_id:     clientId,
                    client_secret: clientSecret,
                    redirect_uri:  redirectUri,
                    grant_type:    'authorization_code'
                }),
                { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
            );
            return response.data;
        } catch (error) {
            console.error("GOOGLE EXCHANGE ERROR:", error.response ? error.response.data : error.message);
            logger.error('Failed to exchange Google code for token', error.response?.data || error.message);
            throw error;
        }
    }

    /**
     * Fetch the authenticated user's Google profile.
     * @param {string} accessToken
     * @returns {Promise<{ sub, email, name, picture }>}
     */
    async getUserProfile(accessToken) {
        try {
            const response = await axios.get(
                'https://www.googleapis.com/oauth2/v3/userinfo',
                { headers: { Authorization: `Bearer ${accessToken}` } }
            );
            return response.data;
        } catch (error) {
            logger.error('Failed to fetch Google user profile', error.response?.data || error.message);
            throw error;
        }
    }
}

module.exports = new GoogleAuthService();
