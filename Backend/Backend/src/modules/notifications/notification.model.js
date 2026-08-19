'use strict';

const { DataTypes } = require('sequelize');

/**
 * Notification Model
 * ----------------------------------------------------------------
 * Per-user in-app notifications (login/pull/refresh success-or-error
 * toasts, etc.), persisted server-side and scoped by userId.
 *
 * Previously these lived client-side in localStorage under the key
 * `fa_notifications` — a single key shared by ANY user who logs into
 * that browser, so User B would see User A's notification history.
 * Storing them here, linked to the owning user's row in `users`, is
 * what makes the history actually per-account.
 * ----------------------------------------------------------------
 */
module.exports = (sequelize) => {
    return sequelize.define(
        'Notification',
        {
            id: {
                type: DataTypes.UUID,
                defaultValue: DataTypes.UUIDV4,
                primaryKey: true
            },

            // FK -> users.id. users.id is a STRING(13) "FIN<year><5 digits>"
            // id (see modules/auth/user.model.js), not an auto-increment
            // integer or a UUID, so this column has to be a matching STRING.
            userId: {
                type: DataTypes.STRING(13),
                allowNull: false,
                references: {
                    model: 'users',
                    key: 'id'
                }
            },

            type: {
                type: DataTypes.ENUM('success', 'error'),
                allowNull: false,
                // Database Validation: mirrors the Joi check in
                // core/validation/schemas.js#createNotification one layer
                // below it, same defense-in-depth pattern as
                // modules/auth/user.model.js#role.
                validate: {
                    isIn: { args: [['success', 'error']], msg: "type must be 'success' or 'error'." }
                }
            },

            message: {
                type: DataTypes.STRING(500),
                allowNull: false
            },

            // Optional technical/secondary detail line shown under the
            // main message (e.g. "Rows 11-20 of 100 written.").
            detail: {
                type: DataTypes.STRING(1000),
                allowNull: true
            },

            // Which ERP connection this notification is about, if any.
            // NULL for provider-agnostic notifications (e.g. "Login
            // successful.") — mirrors NotificationService's existing
            // `_forCurrentContext` provider filtering on the frontend.
            provider: {
                type: DataTypes.ENUM('quickbooks', 'xero'),
                allowNull: true
            },

            read: {
                type: DataTypes.BOOLEAN,
                allowNull: false,
                defaultValue: false
            }
        },
        {
            tableName: 'notifications',
            timestamps: true,
            indexes: [
                // GET /api/notifications always filters by userId and
                // orders by createdAt DESC — see
                // modules/notifications/controller.js#list.
                { fields: ['userId', 'createdAt'] }
            ]
        }
    );
};
