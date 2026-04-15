'use strict';

/**
 * Entry point — Express HTTP server.
 *
 * Endpoints:
 *   POST /api/messages  — Bot Framework webhook (Teams → bot)
 *   GET  /health        — Liveness probe
 */

const express = require('express');
const { BotFrameworkAdapter } = require('botbuilder');

const config = require('./config');
const logger = require('./logger');
const { IcebreakerBot } = require('./bot');

// ---------------------------------------------------------------------------
// Bot Framework adapter
// ---------------------------------------------------------------------------

const adapter = new BotFrameworkAdapter({
  appId: config.microsoftAppId,
  appPassword: config.microsoftAppPassword,
});

adapter.onTurnError = async (context, error) => {
  logger.error(`Unhandled bot error: ${error.message}`, error);
  try {
    await context.sendActivity('Sorry, something went wrong. Please try again.');
  } catch (_) {
    // ignore send failures during error handling
  }
};

const bot = new IcebreakerBot();

// ---------------------------------------------------------------------------
// HTTP server
// ---------------------------------------------------------------------------

const app = express();
app.use(express.json());

app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

app.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

// ---------------------------------------------------------------------------
// Start
// ---------------------------------------------------------------------------

app.listen(config.port, () => {
  logger.info(`Bot listening on port ${config.port}`);
  logger.info(`Messaging endpoint: ${config.callbackBaseUrl}/api/messages`);
});

module.exports = { app, adapter };
