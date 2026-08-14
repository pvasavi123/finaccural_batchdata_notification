const { DataTypes } = require('sequelize');

/**
 * User Model
 * ----------------------------------------------------------------
 * Shared users table used by all authentication providers.
 *
 * provider values: 'local' | 'google' | 'microsoft'
 * role    values: 'user'  | 'admin'  (defaults to 'user')
 * ----------------------------------------------------------------
 */
module.exports = (sequelize) => {
    return sequelize.define(
        'User',
        {
            id: {
                type:          DataTypes.INTEGER,
                primaryKey:    true,
                autoIncrement: true
            },

            name: {
                type:      DataTypes.STRING(100),
                allowNull: false,
                validate: {
                    len: { args: [1, 100], msg: 'Name must be between 1 and 100 characters.' }
                }
            },

            email: {
                type:      DataTypes.STRING(150),
                allowNull: false,
                unique:    { name: 'users_email_unique', msg: 'Email already registered' },
                // Database Validation: enforced again here, one layer below
                // the Joi schema (core/validation/schemas.js) and the regex
                // check in auth.validation.js, so a row can never be
                // written with a malformed email even if it reaches the
                // model through a path that skips those upper layers
                // (e.g. a future admin script or seed).
                validate: {
                    isEmail: { msg: 'A valid email address is required.' }
                }
            },

            // Null for OAuth-only users who never set a local password
            password_hash: {
                type:      DataTypes.STRING,
                allowNull: true
            },

            provider: {
                type:         DataTypes.ENUM('local', 'google', 'microsoft'),
                allowNull:    false,
                defaultValue: 'local'
            },

            // Populated for Google OAuth users
            google_id: {
                type:      DataTypes.STRING,
                allowNull: true
            },

            // Populated for Microsoft Entra ID (Azure AD) OAuth users
            microsoft_id: {
                type:      DataTypes.STRING,
                allowNull: true
            },

            role: {
                type:         DataTypes.STRING(50),
                allowNull:    false,
                defaultValue: 'user',
                validate: {
                    isIn: { args: [['user', 'admin']], msg: "Role must be 'user' or 'admin'." }
                }
            },

            is_active: {
                type:         DataTypes.BOOLEAN,
                allowNull:    false,
                defaultValue: true
            },

            plan: {
                type:      DataTypes.STRING(50),
                allowNull: true
            },

            trial_ends_at: {
                type:      DataTypes.DATE,
                allowNull: true
            }
        },
        {
            tableName:  'users',
            timestamps: true,
            createdAt:  'created_at',
            updatedAt:  'updated_at'
        }
    );
};
