const { createClient } = require('redis');
const config = require('./config');
const logger = require('./logger');
const redisClient = createClient({
    socket: {
        host: config.REDIS.HOST,
        port: config.REDIS.PORT,
        reconnectStrategy: (retries) => {
            logger.info(`Redis reconnect attempt #${retries}`);
            return Math.min(retries * 500, 5000);
        }
    },
    password: config.REDIS.PASSWORD
});
redisClient.on('connect', () => {
    logger.info(`Redis client connecting to ${config.REDIS.HOST}:${config.REDIS.PORT}...`);
});
redisClient.on('ready', () => {
    logger.info('Redis client connected and ready.');
});
redisClient.on('error', (err) => {
    logger.error('Redis client error:', err);
});
redisClient.on('end', () => {
    logger.info('Redis client connection closed.');
});
// Connect asynchronously
redisClient.connect().catch((err) => {
    logger.error('Failed to initialize Redis connection:', err);
});
module.exports = redisClient;
