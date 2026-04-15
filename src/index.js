'use strict';

/**
 * Entry point — Restify HTTP server.
 *
 * Two endpoints:
 *  POST /api/messages   — Bot Framework webhook (Teams → bot)
 *  POST /api/callbacks  — ACS Call Automation event callbacks (ACS → bot)
 *  GET  /health         — Liveness probe (useful for container orchestration)
 */

const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');

const config = require('./config');
const logger = require('./logger');
const callManager = require('./callManager');
const { IcebreakerBot } = require('./bot');

// ---------------------------------------------------------------------------
// Bot Framework adapter
// ---------------------------------------------------------------------------

const adapter = new BotFrameworkAdapter({
  appId: config.microsoftAppId,
  appPassword: config.microsoftAppPassword,
});

// Global error handler for the adapter
adapter.onTurnError = async (context, error) => {
  logger.error(`Unhandled Bot Framework error: ${error.message}`, error);
  await context.sendTraceActivity(
    'OnTurnError',
    `${error}`,
    'https://www.botframework.com/schemas/error',
    'TurnError'
  );
  await context.sendActivity('An error occurred. Please try again.');
};

const bot = new IcebreakerBot();

// ---------------------------------------------------------------------------
// HTTP server
// ---------------------------------------------------------------------------

const server = restify.createServer({ name: 'IcebreakerAudioBot' });
server.use(restify.plugins.bodyParser({ mapParams: false }));

// --- Bot Framework messages (from Teams) ---
server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// --- ACS Call Automation callbacks ---
server.post('/api/callbacks', async (req, res, next) => {
  try {
    // ACS sends an array of CloudEvents
    const events = Array.isArray(req.body) ? req.body : [req.body];

    for (const event of events) {
      await callManager.handleEvent(event);
    }

    res.send(200);
  } catch (err) {
    logger.error(`Error handling ACS callback: ${err.message}`, err);
    res.send(500, { error: err.message });
  }
  return next();
});

// --- Health check ---
server.get('/health', (_req, res, next) => {
  const sessions = callManager.getActiveSessions();
  res.send(200, { status: 'ok', activeSessions: sessions.length });
  return next();
});

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

server.listen(config.port, () => {
  logger.info(`${server.name} listening on port ${config.port}`);
  logger.info(`Bot endpoint : POST ${config.callbackBaseUrl}/api/messages`);
  logger.info(`ACS callback : POST ${config.callbackBaseUrl}/api/callbacks`);
  logger.info(`Health check : GET  ${config.callbackBaseUrl}/health`);
});

module.exports = { server, adapter };
