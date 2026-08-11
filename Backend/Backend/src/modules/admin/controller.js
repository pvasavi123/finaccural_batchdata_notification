'use strict';

const AdminService = require('./service');
const UserRepository = require('../auth/user.repository');
const jwt = require('jsonwebtoken');
const config = require('../../core/config');
const { ValidationError } = require('../../core/errors/AppError');

/**
 * AdminController
 * -----------------------------------------------------------------
 * Handles all incoming HTTP requests for the Admin module.
 * Delegates all business logic to AdminService.
 * The Service already returns clean DTOs via AdminMapper,
 * so the controller can pass them directly to the response.
 * -----------------------------------------------------------------
 */
class AdminController {

    /**
     * POST /api/admin/login
     * Authenticates an admin user.
     */
    async login(req, res, next) {
        try {
            const { email, password } = req.body;

            if (!email || !password) {
                throw new ValidationError('Email and Password are required.');
            }

            const admin = await AdminService.login(email, password);

            req.session.admin = admin;

            const token = jwt.sign(
                { id: admin.id, email: admin.email, role: 'admin' },
                config.JWT_SECRET,
                { expiresIn: '24h' }
            );

            return res.status(200).json({
                success: true,
                message: 'Login Successful',
                admin,
                token
            });
        } catch (error) {
            // Error Validation: let the centralized errorHandler classify
            // anything unexpected (a hardcoded 401 here would previously
            // have mislabeled e.g. a DB outage as "unauthorized").
            next(error);
        }
    }

    /**
     * POST /api/admin/signup
     * Creates a new admin user.
     */
    async signup(req, res, next) {
        try {
            const { name, email, password } = req.body;

            if (!name || !email || !password) {
                throw new ValidationError('Name, Email and Password are required.');
            }

            const admin = await AdminService.signup(name, email, password);

            req.session.admin = admin;

            const token = jwt.sign(
                { id: admin.id, email: admin.email, role: 'admin' },
                config.JWT_SECRET,
                { expiresIn: '24h' }
            );

            return res.status(201).json({
                success: true,
                message: 'Signup Successful',
                admin,
                token
            });
        } catch (error) {
            next(error);
        }
    }

    /**
     * GET /api/admin/users
     * Authorization Validation — admin-only. Lists every FinAccrual user
     * account (id, name, email, provider, role, is_active, plan). Gated
     * by `authenticate JWT` + `authorize('admin')` in routes.js — a
     * regular FinAccrual user's own JWT (role: 'user') is rejected here
     * with a 403 ERR_FORBIDDEN even though it passes JWT verification
     * (it's signed with the same JWT_SECRET as an admin token).
     */
    async listUsers(req, res, next) {
        try {
            const users = await UserRepository.findAll();
            return res.json({ success: true, count: users.length, users });
        } catch (error) {
            next(error);
        }
    }
}

module.exports = new AdminController();