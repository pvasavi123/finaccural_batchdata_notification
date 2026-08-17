'use strict';
const AuthService    = require('./auth.service');
const AuthValidation = require('./auth.validation');
const UserRepository = require('./user.repository');
const OAuthPopupView = require('./views/oauthPopup.view');
const logger         = require('../../core/logger');
const { AppError, ValidationError } = require('../../core/errors/AppError');
/**
 * AuthController
 * ----------------------------------------------------------------
 * Thin HTTP layer — validate input, call AuthService, shape response.
 * No business logic lives here.
 * ----------------------------------------------------------------
 */
class AuthController {
    // ----------------------------------------------------------------
    // POST /api/auth/signup
    // ----------------------------------------------------------------
    async signup(req, res, next) {
        try {
            const { valid, errors } = AuthValidation.validateSignup(req.body);
            if (!valid) {
                throw new ValidationError(errors.join(', '));
            }
            const { name, email, password } = req.body;
            const result = await AuthService.signup(name, email, password);
            return res.status(201).json({
                success: true,
                token:   result.token,
                user:    result.user
            });
        } catch (error) {
            // Database Validation / Error Validation: previously this
            // forced EVERY non-operational error (including a raw DB
            // connection failure like ECONNREFUSED) into a 400
            // ValidationError, which is misleading — "the database is
            // unreachable" is not "your input was invalid". Operational
            // errors (ValidationError, etc.) are passed through as-is;
            // anything else is now handed to the centralized errorHandler
            // (core/middleware/errorHandler.js) so it gets classified
            // properly (503 for a DB/network outage, 500 otherwise)
            // instead of being mislabeled here.
            next(error);
        }
    }
    // ----------------------------------------------------------------
    // POST /api/auth/login
    // ----------------------------------------------------------------
    async login(req, res, next) {
        try {
            const { valid, errors } = AuthValidation.validateLogin(req.body);
            if (!valid) {
                throw new ValidationError(errors.join(', '));
            }
            const { email, password } = req.body;
            const result = await AuthService.login(email, password);
            return res.status(200).json({
                success: true,
                token:   result.token,
                user:    result.user
            });
        } catch (error) {
            // Same fix as signup() above: don't force a raw DB/network
            // error into a misleading 401 "unauthorized" — only genuine
            // bad-credentials failures should look like an auth failure,
            // and AuthService.login already throws an operational error
            // for that case. Let errorHandler.js classify anything else.
            next(error);
        }
    }
    // ----------------------------------------------------------------
    // GET /api/auth/google/connect?login_hint=<email>
    // login_hint is optional — passed by the account picker when the
    // user clicked a remembered "previous account" row, so Google can
    // skip straight past its account-chooser screen. Left off, this
    // behaves exactly as before (fresh "Add account" picks).
    // ----------------------------------------------------------------
    googleConnect(req, res) {
        try {
            const loginHint = AuthController._safeLoginHint(req.query.login_hint);
            const authUrl = AuthService.getGoogleAuthUrl(loginHint);
            res.redirect(authUrl);
        } catch (error) {
            logger.error('Error generating Google OAuth URL', error);
            res.status(500).json({ success: false, message: 'Failed to generate authorisation URL' });
        }
    }
    // ----------------------------------------------------------------
    // GET /api/auth/google/callback
    // ----------------------------------------------------------------
    async googleCallback(req, res) {
        try {
            const { code } = req.query;
            if (!code) throw new Error('No authorisation code provided');
            const { user, token } = await AuthService.handleGoogleCallback(code);
            return res.send(OAuthPopupView.renderAuthSuccess({ provider: 'google', email: user.email, name: user.name, token }));
        } catch (error) {
            return res.send(OAuthPopupView.renderError({ provider: 'google', message: error.message }));
        }
    }
    // ----------------------------------------------------------------
    // GET /api/microsoft/connect?login_hint=<email>  (aliased at /api/auth/microsoft/connect)
    // Same login_hint contract as googleConnect() above.
    // ----------------------------------------------------------------
    microsoftConnect(req, res) {
        try {
            const loginHint = AuthController._safeLoginHint(req.query.login_hint);
            const authUrl = AuthService.getMicrosoftAuthUrl(loginHint);
            res.redirect(authUrl);
        } catch (error) {
            logger.error('Error generating Microsoft OAuth URL', error);
            res.status(500).json({ success: false, message: 'Failed to generate authorisation URL' });
        }
    }
    // ----------------------------------------------------------------
    // GET /api/microsoft/callback  (aliased at /api/auth/microsoft/callback)
    // ----------------------------------------------------------------
    async microsoftCallback(req, res) {
        try {
            const { code } = req.query;
            if (!code) throw new Error('No authorisation code provided');
            const { user, token } = await AuthService.handleMicrosoftCallback(code);
            return res.send(OAuthPopupView.renderAuthSuccess({ provider: 'microsoft', email: user.email, name: user.name, token }));
        } catch (error) {
            return res.send(OAuthPopupView.renderError({ provider: 'microsoft', message: error.message }));
        }
    }
    // ----------------------------------------------------------------
    // GET /api/auth/me   (protected)
    // ----------------------------------------------------------------
    async getMe(req, res, next) {
        try {
            const user = await UserRepository.findById(req.user.userId);
            if (!user) {
                throw new AppError('The requested resource was not found.', 404, 'ERR_NOT_FOUND', 'User not found.');
            }
            return res.json({
                success: true,
                user: {
                    id:             user.id,
                    name:           user.name,
                    email:          user.email,
                    role:           user.role,
                    provider:       user.provider,
                    plan:           user.plan,
                    trialEndsAt:    user.trial_ends_at,
                    subscriptionId: user.id
                }
            });
        } catch (error) {
            next(error instanceof Error && error.isOperational ? error : new AppError('Something went wrong on our end. Please try again later.', 500, 'ERR_INTERNAL', error.message));
        }
    }
    // ----------------------------------------------------------------
    // POST /api/auth/logout
    // ----------------------------------------------------------------
    // JWT auth is stateless — the frontend already discards its token and
    // redirects to the Login view on its own. This endpoint just clears
    // any server-side OAuth session leftovers (req.session.user_mail,
    // xero_pending_*, etc.) so a stale session cookie can't leak into the
    // next sign-in attempt on the same browser.
    async logout(req, res, next) {
        try {
            if (req.session) {
                req.session.destroy(() => {});
            }
            return res.json({ success: true, message: 'Logged out.' });
        } catch (error) {
            next(error instanceof Error && error.isOperational ? error : new AppError('Something went wrong on our end. Please try again later.', 500, 'ERR_INTERNAL', error.message));
        }
    }
    // ----------------------------------------------------------------
    // POST /api/auth/update-plan
    // ----------------------------------------------------------------
    async updatePlan(req, res, next) {
        try {
            const { plan } = req.body;
            if (!plan) {
                throw new ValidationError('Plan is required.');
            }
            const user = await UserRepository.findById(req.user.userId);
            const oldPlan = user ? user.plan : null;
            const planWeights = {
                'basic': 1,
                'standard': 2,
                'pro': 3
            };
            // A user with no prior plan (brand-new signup) is never
            // "downgrading" — defaulting oldPlanKey to 'Pro' here would
            // treat their very first plan selection (Basic/Standard) as a
            // downgrade from Pro and wipe out connections they just made.
            const oldPlanKey = oldPlan ? oldPlan.toLowerCase() : null;
            const newPlanKey = plan.toLowerCase();
            const isDowngrade = oldPlanKey !== null && (planWeights[newPlanKey] || 0) < (planWeights[oldPlanKey] || 0);
            await UserRepository.update(req.user.userId, { plan });
            if (isDowngrade && req.user.email) {
                const { QuickBooksToken, XeroToken } = require('../../core/database');
                await QuickBooksToken.destroy({ where: { mail: req.user.email } });
                await XeroToken.destroy({ where: { mail: req.user.email } });
                logger.info(`User ${req.user.userId} downgraded from ${oldPlan} to ${plan}. Cleared company connections.`);
            }
            return res.json({ success: true, message: 'Plan updated successfully' });
        } catch (error) {
            next(error instanceof Error && error.isOperational ? error : new AppError('Something went wrong on our end. Please try again later.', 500, 'ERR_INTERNAL', error.message));
        }
    }
    // ----------------------------------------------------------------
    // POST /api/auth/start-trial
    // Explicit "Start Free Trial" choice from the Free Trial vs
    // Subscription Plan screen — sets plan + a fresh trial_ends_at
    // clock starting now. New accounts get no plan at signup (see
    // AuthService._newAccountDefaults()), so this is the only place
    // the trial actually begins.
    // ----------------------------------------------------------------
    async startTrial(req, res, next) {
        try {
            const { user } = await AuthService.startTrial(req.user.userId);
            return res.json({ success: true, user });
        } catch (error) {
            next(error instanceof Error && error.isOperational ? error : new AppError('Something went wrong on our end. Please try again later.', 500, 'ERR_INTERNAL', error.message));
        }
    }

    // ----------------------------------------------------------------
    // Helpers
    // ----------------------------------------------------------------

    /**
     * Validates a `login_hint` query param before it's forwarded to
     * Google/Microsoft's authorisation endpoint. Deliberately strict —
     * this only ever needs to carry an email address the account picker
     * already showed the user, so anything that doesn't look like one
     * (wrong type, too long, no '@') is dropped rather than forwarded.
     * @param {*} raw
     * @returns {string|undefined}
     */
    static _safeLoginHint(raw) {
        if (typeof raw !== 'string') return undefined;
        const trimmed = raw.trim();
        if (!trimmed || trimmed.length > 254 || !trimmed.includes('@')) return undefined;
        return trimmed;
    }
}
module.exports = new AuthController();
