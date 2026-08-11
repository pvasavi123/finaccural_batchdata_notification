const { DataTypes } = require('sequelize');

module.exports = (sequelize) => {
    return sequelize.define('QuickBooksToken', {
        realm_id: {
            type: DataTypes.STRING(50),
            primaryKey: true,
            allowNull: false
        },
        access_token: {
            type: DataTypes.TEXT,
            allowNull: false
        },
        refresh_token: {
            type: DataTypes.TEXT,
            allowNull: false
        },
        token_type: {
            type: DataTypes.STRING(50)
        },
        expires_in: {
            type: DataTypes.INTEGER,
            defaultValue: 0
        },
        x_refresh_token_expires_in: {
            type: DataTypes.INTEGER,
            defaultValue: 0
        },
        session_info: {
            type: DataTypes.TEXT,
            allowNull: true
        },
        mail: {
            type: DataTypes.STRING(255),
            allowNull: true
        },
        company_name: {
            type: DataTypes.STRING(255),
            allowNull: true,
            defaultValue: 'QuickBooks Company'
        },
        status: {
            // 'Not Synced' until the first successful Master Data Pull sets
            // it to 'Active' (see QuickBooksService.pullMasterData);
            // 'Disconnected' once the user disconnects or a refresh fails.
            type: DataTypes.STRING(20),
            defaultValue: 'Not Synced',
            // Database Validation: guards against a future typo (e.g.
            // 'Actve') silently writing an unrecognized status that the
            // rest of the app (repository.js WHERE status: 'Active'
            // filters) would then just quietly never match.
            validate: {
                isIn: {
                    args: [['Not Synced', 'Active', 'Disconnected']],
                    msg: "status must be one of 'Not Synced', 'Active', 'Disconnected'."
                }
            }
        },
        last_synced_at: {
            type: DataTypes.DATE,
            allowNull: true
        }
    }, {
        tableName: 'quickbooks_quickbookstoken',
        timestamps: true,
        createdAt: 'created_at',
        updatedAt: 'updated_at'
    });
};
