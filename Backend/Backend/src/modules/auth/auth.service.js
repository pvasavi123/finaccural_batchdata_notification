'use strict';

const bcrypt           = require('bcrypt');
const UserRepository   = require('./user.repository');
const JwtService       = require('./jwt.service');
const GoogleService    = require('./google.service');
const MicrosoftService = require('./microsoft.service');
const logger           = require('../../core/logger');
const config           = require('../../core/config');
const { ValidationError, AuthenticationError } = require('../../core/errors/AppError');

/**
 * AuthService
 * ----------------------------------------------------------------
 * Central business-logic layer for all authentication flows:
 *   - Local signup / login
 *   - Google OAuth callback (upsert user, return JWT)
 *
 * Controllers stay thin — they only call methods here and shape
 * the HTTP response.
 * ----------------------------------------------------------------
 */
class AuthService {

    // ----------------------------------------------------------------
    // Helpers
    // ----------------------------------------------------------------

    /**
     * Build a safe public DTO from a User model instance.
     * Never exposes password_hash or google_id.
     * @param {object} user - Sequelize User instance
     * @returns {{ id, name, email, role, provider }}
     */
    static _toUserDTO(user) {
        return {
            id:       user.id,
            name:     user.name,
            email:    user.email,
            role:     user.role,
            provider: user.provider,
            plan:     user.plan,
            trialEndsAt: user.trial_ends_at
        };
    }

    /**
     * Fields every brand-new signup gets, regardless of provider.
     * Deliberately leaves plan/trial_ends_at unset — auto-enrolling a
     * brand-new account into the trial here meant the frontend's "no
     * plan yet" check (checkSubscription -> !!user.plan) was never
     * false, so new users skipped straight past the Free Trial vs
     * Subscription Plan screen (AppController.openTrialSelectDialog)
     * into an already-running trial they never chose. The trial clock
     * now only starts once the user explicitly picks "Start Free
     * Trial" — see startTrial() below.
     * Centralised here so signup(), handleGoogleCallback(), and
     * handleMicrosoftCallback() can't drift out of sync on this.
     * @returns {{ plan: null, trial_ends_at: null }}
     */
    static _newAccountDefaults() {
        return {
            plan:          null,
            trial_ends_at: null
        };
    }

    /**
     * Starts the free trial for an already-authenticated user — called
     * when they explicitly choose "Start Free Trial" on the Free Trial
     * vs Subscription Plan screen. Sets the plan and a fresh
     * trial_ends_at clock starting now (config.TRIAL.DURATION_MS long).
     * @param {number} userId
     * @returns {Promise<{ user: UserDTO }>}
     */
    static async startTrial(userId) {
        const user = await UserRepository.update(userId, {
            plan:          config.TRIAL.DEFAULT_PLAN,
            trial_ends_at: new Date(Date.now() + config.TRIAL.DURATION_MS)
        });
        return { user: AuthService._toUserDTO(user) };
    }

    /**
     * Generate a JWT whose payload matches the spec in the plan.
     * @param {object} user - Sequelize User instance
     * @returns {string} signed JWT
     */
    static _buildToken(user) {
        return JwtService.generateToken({
            userId: user.id,
            email:  user.email,
            role:   user.role
        });
    }

    // ----------------------------------------------------------------
    // Local Email / Password
    // ----------------------------------------------------------------

    /**
     * Register a new local (email + password) user.
     *
     * @param {string} name
     * @param {string} email
     * @param {string} password   Plain-text; hashed here before storage.
     * @returns {Promise<{ token: string, user: UserDTO }>}
     * @throws {Error} if email already registered
     */
    static async signup(name, email, password) {
        const normalised = email.toLowerCase().trim();

        // Database Validation (service-level, pre-write): check uniqueness
        // ourselves and fail with a proper 400 ValidationError, rather than
        // letting the insert hit the DB's unique constraint and bubble up
        // as a raw SequelizeUniqueConstraintError that errorHandler.js
        // would otherwise have to guess at classifying.
        const existing = await UserRepository.findByEmail(normalised);
        if (existing) {
            throw new ValidationError('Email already registered.');
        }

        const password_hash = await bcrypt.hash(password, 10);

        const user = await UserRepository.create({
            name:          (name || '').trim() || 'FinAccrual User',
            email:         normalised,
            password_hash,
            provider:      'local',
            role:          'user',
            ...AuthService._newAccountDefaults()
        });

        return {
            token: AuthService._buildToken(user),
            user:  AuthService._toUserDTO(user)
        };
    }

    /**
     * Authenticate a local user with email and password.
     *
     * @param {string} email
     * @param {string} password   Plain-text password to check.
     * @returns {Promise<{ token: string, user: UserDTO }>}
     * @throws {Error} if credentials are invalid
     */
    static async login(email, password) {
        const normalised = email.toLowerCase().trim();
        const user       = await UserRepository.findByEmail(normalised);

        // Authentication Validation: each of these three checks throws a
        // proper operational 401 (AuthenticationError) instead of a plain
        // Error, so it's classified correctly by errorHandler.js and by
        // auth.controller.js's `next(error)` — a real "unreachable
        // database" failure now stays distinguishable from "wrong
        // password" (see auth.controller.js login() for the other half
        // of this fix).
        if (!user) {
            throw new AuthenticationError('Invalid email or password.');
        }

        if (user.provider !== 'local' || !user.password_hash) {
            throw new AuthenticationError('This account uses Google sign-in. Please use "Continue with Google".');
        }

        const isMatch = await bcrypt.compare(password, user.password_hash);
        if (!isMatch) {
            throw new AuthenticationError('Invalid email or password.');
        }

        return {
            token: AuthService._buildToken(user),
            user:  AuthService._toUserDTO(user)
        };
    }

    // ----------------------------------------------------------------
    // Google OAuth
    // ----------------------------------------------------------------

    /**
     * Return the Google OAuth 2.0 authorisation URL to redirect the
     * browser to.
     * @param {string} [loginHint] - Previously-seen account's email; see
     *   GoogleService.getAuthUrl() for why this skips the account chooser.
     * @returns {string}
     */
    static getGoogleAuthUrl(loginHint) {
        return GoogleService.getAuthUrl(loginHint);
    }

    /**
     * Handle the Google OAuth callback.
     * - Exchanges the code for tokens.
     * - Fetches the Google profile.
     * - Upserts the user in the database (create if new, update google_id if returning).
     * - Returns the user DTO and JWT.
     *
     * @param {string} code   OAuth2 authorisation code from query string.
     * @returns {Promise<{ token: string, user: UserDTO, isNewUser: boolean }>}
     */
    static async handleGoogleCallback(code) {
        const tokens  = await GoogleService.exchangeCodeForToken(code);
        const profile = await GoogleService.getUserProfile(tokens.access_token);

        const googleId = profile.sub;
        const email    = (profile.email || '').toLowerCase().trim();
        const name     = profile.name  || profile.email || 'User';

        // 1. Try to find by Google ID (fastest, most stable)
        let user = await UserRepository.findByGoogleId(googleId);

        // 2. Fall back to email lookup (handles users who signed up locally first)
        if (!user) {
            user = await UserRepository.findByEmail(email);
        }

        if (user) {
            // Returning user — ensure google_id is persisted if missing
            if (!user.google_id) {
                user = await UserRepository.update(user.id, {
                    google_id: googleId,
                    provider:  'google'
                });
            }
        } else {
            // New user — create account
            user = await UserRepository.create({
                name,
                email,
                provider:  'google',
                google_id: googleId,
                role:      'user',
                ...AuthService._newAccountDefaults()
            });
        }

        const isNewUser = !user.created_at ||
            (new Date() - new Date(user.created_at)) < 5000;

        return {
            token:     AuthService._buildToken(user),
            user:      AuthService._toUserDTO(user),
            isNewUser: user.provider === 'google' && !user.password_hash
        };
    }

    // ----------------------------------------------------------------
    // Microsoft Entra ID (Azure AD) OAuth
    // ----------------------------------------------------------------

    /**
     * Return the Microsoft Entra ID OAuth 2.0 authorisation URL to
     * redirect the browser to.
     * @param {string} [loginHint] - Previously-seen account's email; see
     *   MicrosoftService.getAuthUrl() for why this skips the account
     *   chooser.
     * @returns {string}
     */
    static getMicrosoftAuthUrl(loginHint) {
        return MicrosoftService.getAuthUrl(loginHint);
    }

    /**
     * Handle the Microsoft Entra ID OAuth callback.
     * - Exchanges the code for tokens.
     * - Fetches the Microsoft Graph profile.
     * - Upserts the user in the database (create if new, update
     *   microsoft_id if returning).
     * - Returns the user DTO and JWT.
     *
     * @param {string} code   OAuth2 authorisation code from query string.
     * @returns {Promise<{ token: string, user: UserDTO, isNewUser: boolean }>}
     */
    static async handleMicrosoftCallback(code) {
        const tokens  = await MicrosoftService.exchangeCodeForToken(code);
        const profile = await MicrosoftService.getUserProfile(tokens.access_token);

        const microsoftId = profile.sub;
        const email        = (profile.email || '').toLowerCase().trim();
        const name          = profile.name || profile.email || 'User';

        if (!email) {
            throw new Error('Microsoft account has no email or userPrincipalName to sign in with');
        }

        // 1. Try to find by Microsoft ID (fastest, most stable)
        let user = await UserRepository.findByMicrosoftId(microsoftId);

        // 2. Fall back to email lookup (handles users who signed up locally
        //    or with Google first)
        if (!user) {
            user = await UserRepository.findByEmail(email);
        }

        if (user) {
            // Returning user — ensure microsoft_id is persisted if missing
            if (!user.microsoft_id) {
                user = await UserRepository.update(user.id, {
                    microsoft_id: microsoftId,
                    provider:     'microsoft'
                });
            }
        } else {
            // New user — create account
            user = await UserRepository.create({
                name,
                email,
                provider:     'microsoft',
                microsoft_id: microsoftId,
                role:         'user',
                ...AuthService._newAccountDefaults()
            });
        }

        const isNewUser = !user.created_at ||
            (new Date() - new Date(user.created_at)) < 5000;

        return {
            token:     AuthService._buildToken(user),
            user:      AuthService._toUserDTO(user),
            isNewUser: user.provider === 'microsoft' && !user.password_hash
        };
    }
}

module.exports = AuthService;
