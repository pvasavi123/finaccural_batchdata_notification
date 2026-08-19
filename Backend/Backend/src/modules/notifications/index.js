'use strict';

/**
 * Notifications Module Entry Point
 * -----------------------------------------------------------------
 * Exports the public API of this module so that the rest of the
 * application only depends on this file, not on internal details.
 * -----------------------------------------------------------------
 */
const routes = require('./routes');
const model  = require('./notification.model');

module.exports = { routes, model };
