'use strict';

/**
 * Excel Validation Module Entry Point
 * -----------------------------------------------------------------
 * Exports the public API of this module so that the rest of the
 * application only depends on this file, not on internal details.
 * -----------------------------------------------------------------
 */
const routes  = require('./routes');
const service = require('./service');
const schemas = require('./schemas/masterDataSchemas');

module.exports = { routes, service, schemas };
