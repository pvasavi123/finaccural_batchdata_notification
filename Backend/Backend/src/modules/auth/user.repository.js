'use strict';

const { User } = require('../../core/database');

/**
 * UserRepository
 * ----------------------------------------------------------------
 * Data-access layer for the users table.
 * All Sequelize queries are contained here; services never import
 * models directly.
 * ----------------------------------------------------------------
 */
class UserRepository {

    /**
     * Find a user by their email address (case-insensitive via DB collation).
     * @param {string} email
     * @returns {Promise<User|null>}
     */
    static async findByEmail(email) {
        return await User.findOne({ where: { email: email.toLowerCase().trim() } });
    }

    /**
     * Find a user by their primary key.
     * @param {number} id
     * @returns {Promise<User|null>}
     */
    static async findById(id) {
        return await User.findByPk(id);
    }

    /**
     * Find a user by their Google OAuth subject ID.
     * @param {string} googleId
     * @returns {Promise<User|null>}
     */
    static async findByGoogleId(googleId) {
        return await User.findOne({ where: { google_id: googleId } });
    }

    /**
     * Find a user by their Microsoft Entra ID (Azure AD) subject ID.
     * @param {string} microsoftId
     * @returns {Promise<User|null>}
     */
    static async findByMicrosoftId(microsoftId) {
        return await User.findOne({ where: { microsoft_id: microsoftId } });
    }

    /**
     * Create a new user record.
     * @param {{ name, email, password_hash?, provider, google_id?, microsoft_id?, role? }} data
     * @returns {Promise<User>}
     */
    static async create(data) {
        return await User.create(data);
    }

    /**
     * Update fields on an existing user.
     * @param {number} id
     * @param {object} data
     * @returns {Promise<User>}
     */
    static async update(id, data) {
        await User.update(data, { where: { id } });
        return await User.findByPk(id);
    }

    /**
     * List every FinAccrual user account. Admin-only — see
     * modules/admin/controller.js listUsers(), gated by
     * core/middleware/authorize.js.
     * @returns {Promise<User[]>}
     */
    static async findAll() {
        return await User.findAll({
            attributes: ['id', 'name', 'email', 'provider', 'role', 'is_active', 'plan', 'trial_ends_at', 'created_at'],
            order: [['created_at', 'DESC']]
        });
    }
}

module.exports = UserRepository;
