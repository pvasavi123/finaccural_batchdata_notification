const dns = require('dns');


if (typeof dns.setDefaultResultOrder === 'function') {
    dns.setDefaultResultOrder('ipv4first');
}

const app = require('./app');
const { sequelize } = require('./core/database');
const config = require('./core/config');
const logger = require('./core/logger');

const shouldAlter = process.env.DB_SYNC_ALTER === 'true';
if (shouldAlter) {
    logger.info('DB_SYNC_ALTER=true — syncing with { alter: true }. Turn this off once your schema is stable.');
}

sequelize.sync({ alter: shouldAlter }).then(() => {
    logger.info("Database synchronized.");
    app.listen(config.PORT, () => {
        logger.info(`Node.js backend running on port ${config.PORT}`);
    });
}).catch(err => {
    logger.error("Unable to connect to the database:", err);
});
