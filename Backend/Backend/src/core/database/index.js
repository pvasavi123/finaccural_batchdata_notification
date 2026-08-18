const { Sequelize } = require('sequelize');
const config = require('../config');

const sequelize = new Sequelize(config.DB.NAME, config.DB.USER, config.DB.PASSWORD, {
    host: config.DB.HOST,
    port: config.DB.PORT,
    dialect: 'mysql',
    logging: false
});

const QuickBooksToken = require('../../modules/quickbooks/model')(sequelize);
const XeroToken       = require('../../modules/xero/model')(sequelize);

// Unified users table used by the auth module
const User = require('../../modules/auth/user.model')(sequelize);

// Backward-compatible alias — the legacy admin module still imports
// { Admin } from here. Pointing it to User means the admins table and
// users table share the same model during the transition period.
// Once admin/repository.js and admin/service.js are removed, this alias
// can be deleted.
const Admin = require('../../modules/admin/model')(sequelize);

module.exports = {
    sequelize,
    QuickBooksToken,
    XeroToken,
    User,
    Admin   // legacy alias — will be removed after admin module is retired
};
