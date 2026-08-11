const { DataTypes } = require("sequelize");

module.exports = (sequelize) => {
  return sequelize.define(
    "XeroToken",
    {
      tenant_id: {
        type: DataTypes.STRING(255),
        primaryKey: true,
        allowNull: false,
      },
      access_token: {
        type: DataTypes.TEXT,
        allowNull: false,
      },
      refresh_token: {
        type: DataTypes.TEXT,
        allowNull: false,
      },
      expires_in: {
        type: DataTypes.INTEGER,
        allowNull: false,
      },
      token_type: {
        type: DataTypes.STRING,
        allowNull: false,
      },
      scope: {
        type: DataTypes.TEXT,
        allowNull: true,
      },
      session_info: {
        type: DataTypes.TEXT,
        allowNull: true,
      },
      mail: {
        type: DataTypes.STRING(255),
        allowNull: true,
      },
      company_name: {
        type: DataTypes.STRING(255),
        allowNull: true,
        defaultValue: 'Xero Organisation'
      },
      status: {
        // 'Not Synced' until the first successful Master Data Pull sets it
        // to 'Active' (see XeroService.pullMasterData); 'Disconnected' once
        // the user disconnects or a refresh fails.
        type: DataTypes.STRING(20),
        defaultValue: 'Not Synced',
        // Database Validation: same guard as QuickBooksToken.status —
        // rejects an unrecognized status before it's written.
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
      },
    },
    {
      tableName: "xero_tokens",
      underscored: true,
      timestamps: true,
    }
  );
};