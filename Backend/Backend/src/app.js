const express = require('express');
const cors = require('cors');

const routes = require('./routes');
const config = require('./core/config');
const { errorHandler, notFoundHandler } = require('./core/middleware/errorHandler');

const path = require("path");

const app = express();

app.use(cors());
app.use(express.json());

const session = require('express-session');
app.use(session({
    secret: config.SESSION_SECRET || 'finaccrual-fallback-secret-key',
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: process.env.NODE_ENV === 'production',
        maxAge: 1000 * 60 * 60 * 24 // 24 hours
    }
}));



app.get('/favicon.ico', (req, res) => res.status(204).end());

app.use('/api', routes);


app.use(notFoundHandler);
app.use(errorHandler);


module.exports = app;