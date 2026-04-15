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
        'Copy .env.example to .env and fill in all required values.\n' +
        'See README.md → Quick Start for the minimal set of variables.'
    );
  }
  return value;
}

// ---------------------------------------------------------------------------
// Operating mode
// ---------------------------------------------------------------------------

/**
 * AUDIO_MODE controls whether the bot joins meetings as a voice participant.
 *
 * false (default) — Chat mode: minimal infrastructure.
 *   Required:  MICROSOFT_APP_ID, MICROSOFT_APP_PASSWORD, CALLBACK_BASE_URL
 *   Optional:  OPENAI_API_KEY or AZURE_OPENAI_* for AI-generated responses
 *
 * true — Audio mode: bot speaks and listens via Azure Communication Services.
 *   Required:  everything above + ACS_CONNECTION_STRING + COGNITIVE_SERVICES_ENDPOINT
 */
const audioMode = process.env.AUDIO_MODE === 'true';

// ---------------------------------------------------------------------------
// Config object
// ---------------------------------------------------------------------------

const config = {
  // Operating mode
  audioMode,

  // HTTP server port
  port: parseInt(process.env.PORT || '3978', 10),

  // -------------------------------------------------------------------------
  // Bot Framework credentials (Azure AD App Registration)
  // Required in both modes.
  // -------------------------------------------------------------------------
  microsoftAppId: requireEnv('MICROSOFT_APP_ID'),
  microsoftAppPassword: requireEnv('MICROSOFT_APP_PASSWORD'),

  // -------------------------------------------------------------------------
  // Public base URL of this bot
  // ACS / Teams sends events to <callbackBaseUrl>/api/callbacks
  // Teams sends messages to    <callbackBaseUrl>/api/messages
  //
  // Local dev: use ngrok -> https://ngrok.com/download
  //   ngrok http 3978   ->   CALLBACK_BASE_URL=https://xxx.ngrok-free.app
  // -------------------------------------------------------------------------
  callbackBaseUrl: requireEnv('CALLBACK_BASE_URL').replace(/\/$/, ''),

  // -------------------------------------------------------------------------
  // Azure Communication Services (required only when AUDIO_MODE=true)
  // -------------------------------------------------------------------------
  acsConnectionString: audioMode
    ? requireEnv('ACS_CONNECTION_STRING')
    : (process.env.ACS_CONNECTION_STRING || ''),

  // Azure Cognitive Services endpoint (enables TTS + STT inside ACS calls)
  cognitiveServicesEndpoint: audioMode
    ? requireEnv('COGNITIVE_SERVICES_ENDPOINT').replace(/\/$/, '')
    : (process.env.COGNITIVE_SERVICES_ENDPOINT || '').replace(/\/$/, ''),

  // -------------------------------------------------------------------------
  // AI provider (optional — falls back to built-in scripted responses)
  // Azure OpenAI takes precedence over standard OpenAI when both are set.
  // -------------------------------------------------------------------------

  // Azure OpenAI
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT || null,
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY || null,
  azureOpenAIDeployment: process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o',
  azureOpenAIApiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-10-21',

  // Standard OpenAI
  openAIKey: process.env.OPENAI_API_KEY || null,
  openAIModel: process.env.OPENAI_MODEL || 'gpt-4o-mini',

  // -------------------------------------------------------------------------
  // Icebreaker activity settings
  // -------------------------------------------------------------------------

  // Default activity type. Leave blank for random selection each session.
  // Options: two_truths_and_a_lie | would_you_rather | fun_fact_round | speed_intro
  defaultActivity: process.env.DEFAULT_ACTIVITY || null,

  // Voice used for TTS in audio mode (any Azure Neural voice name)
  botVoice: process.env.BOT_VOICE || 'en-US-JennyNeural',

  // Milliseconds of trailing silence before the bot considers a response finished
  recognizeSilenceTimeoutMs: parseInt(process.env.RECOGNIZE_SILENCE_TIMEOUT_MS || '4000', 10),

  // Milliseconds to wait for a participant to start speaking before timing out
  recognizeInitialSilenceTimeoutMs: parseInt(
    process.env.RECOGNIZE_INITIAL_SILENCE_TIMEOUT_MS || '30000',
    10
  ),

  // Log level: error | warn | info | debug
  logLevel: process.env.LOG_LEVEL || 'info',
};

module.exports = config;
