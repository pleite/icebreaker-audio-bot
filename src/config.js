'use strict';

require('dotenv').config();

/**
 * Reads and validates required environment variables.
 * Throws a descriptive error on startup if any required variable is missing.
 */
function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(
      `Missing required environment variable: ${name}\n` +
        'Copy .env.example to .env and fill in all required values.'
    );
  }
  return value;
}

const config = {
  // HTTP server port
  port: parseInt(process.env.PORT || '3978', 10),

  // Bot Framework credentials (from Azure Bot Service)
  microsoftAppId: requireEnv('MICROSOFT_APP_ID'),
  microsoftAppPassword: requireEnv('MICROSOFT_APP_PASSWORD'),

  // Azure Communication Services (ACS)
  acsConnectionString: requireEnv('ACS_CONNECTION_STRING'),

  // Public base URL of this bot (used by ACS to send event callbacks)
  // Use ngrok for local development: https://ngrok.com
  callbackBaseUrl: requireEnv('CALLBACK_BASE_URL').replace(/\/$/, ''),

  // Azure Cognitive Services endpoint (enables TTS + STT inside ACS calls)
  cognitiveServicesEndpoint: requireEnv('COGNITIVE_SERVICES_ENDPOINT').replace(/\/$/, ''),

  // Optional: default icebreaker activity to run
  // Options: 'two_truths_and_a_lie' | 'would_you_rather' | 'fun_fact_round' | 'speed_intro'
  // Leave blank for random selection each call
  defaultActivity: process.env.DEFAULT_ACTIVITY || null,

  // Voice used for TTS (any Azure Neural voice name)
  botVoice: process.env.BOT_VOICE || 'en-US-JennyNeural',

  // Milliseconds of silence before the bot considers a response complete
  recognizeSilenceTimeoutMs: parseInt(process.env.RECOGNIZE_SILENCE_TIMEOUT_MS || '4000', 10),

  // Milliseconds to wait for a participant to start speaking
  recognizeInitialSilenceTimeoutMs: parseInt(
    process.env.RECOGNIZE_INITIAL_SILENCE_TIMEOUT_MS || '30000',
    10
  ),

  // Log level: error | warn | info | debug
  logLevel: process.env.LOG_LEVEL || 'info',
};

module.exports = config;
