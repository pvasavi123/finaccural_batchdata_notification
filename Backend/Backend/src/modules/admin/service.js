'use strict';

const bcrypt          = require('bcrypt');
const AdminRepository = require('./repository');
const AdminMapper     = require('./mapper');
const { AuthenticationError, ValidationError } = require('../../core/errors/AppError');

/**
 * AdminService
 * -----------------------------------------------------------------
 * Responsible for all Admin business logic.
 * Uses AdminMapper to strip sensitive fields before returning data.
 * -----------------------------------------------------------------
 */
class AdminService {

    /**
     * Authenticate an admin by email and password.
     * @param {string} email
     * @param {string} password
     * @returns {AdminDTO} clean DTO (no password field)
     */
    static async login(email, password) {
        const admin = await AdminRepository.findByEmail(email);

        // Authentication Validation: proper operational 401s (same fix as
        // modules/auth/auth.service.js login()) instead of plain Error,
        // so a DB outage here is still correctly classified as 503 by
        // errorHandler.js rather than being mislabeled as bad credentials.
        if (!admin) {
            throw new AuthenticationError('Invalid email or password.');
        }

        const isMatch = await bcrypt.compare(password, admin.password);

        if (!isMatch) {
            throw new AuthenticationError('Invalid email or password.');
        }

        return AdminMapper.toAdminDTO(admin);
    }

    /**
     * Create a new admin account.
     * @param {string} name
     * @param {string} email
     * @param {string} password
     * @returns {AdminDTO} clean DTO
     */
    static async signup(name, email, password) {
        const existing = await AdminRepository.findByEmail(email);
        if (existing) {
            throw new ValidationError('Email already registered.');
        }

        const hashedPassword = await bcrypt.hash(password, 10);
        const admin = await AdminRepository.create({
            name,
            email,
            password: hashedPassword
        });

        return AdminMapper.toAdminDTO(admin);
    }
}

module.exports = AdminService;