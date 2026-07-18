/**
 * Vercel serverless entrypoint wrapping the Express application.
 */
const app = require('../server/app');
module.exports = app;
