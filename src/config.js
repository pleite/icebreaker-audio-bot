'use strict';

require('dotenv').config();

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(
      `Missing required environment variable: ${name}\n` +
        'Copy .env.example to .env and fill in all values.\n' +
        'See README.md for the Quick Start guide.'
    );
  }
  return value;
}

const config = {
  // -------------------------------------------------------------------------
  // Server
  // -------------------------------------------------------------------------
  port: parseInt(process.env.PORT || '3978', 10),

  // -------------------------------------------------------------------------
  // Bot Framework — Azure AD App Registration (required)
  // Get these from: Azure Portal → App Registrations → your bot app
  // -------------------------------------------------------------------------
  microsoftAppId: requireEnv('MICROSOFT_APP_ID'),
  microsoftAppPassword: requireEnv('MICROSOFT_APP_PASSWORD'),

  // -------------------------------------------------------------------------
  // Public URL where Teams sends webhook events (required)
  // Local dev: start ngrok first → `ngrok http 3978`
  //            then set this to the HTTPS forwarding URL
  // Production: your Azure Web App or Container App URL
  // -------------------------------------------------------------------------
  callbackBaseUrl: requireEnv('CALLBACK_BASE_URL').replace(/\/$/, ''),

  // -------------------------------------------------------------------------
  // AI provider — optional, graceful fallback to scripted responses
  //
  // Option A: Azure OpenAI  (set AZURE_OPENAI_ENDPOINT + AZURE_OPENAI_API_KEY)
  // Option B: OpenAI        (set OPENAI_API_KEY)
  // -------------------------------------------------------------------------
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT || null,
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY || null,
  azureOpenAIDeployment: process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o',
  azureOpenAIApiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-10-21',
  openAIKey: process.env.OPENAI_API_KEY || null,
  openAIModel: process.env.OPENAI_MODEL || 'gpt-4o-mini',

  // -------------------------------------------------------------------------
  // Icebreaker settings — all optional
  // -------------------------------------------------------------------------

  // Force a specific activity; leave blank for random selection each session.
  // Values: two_truths_and_a_lie | would_you_rather | fun_fact_round | speed_intro
  defaultActivity: process.env.DEFAULT_ACTIVITY || null,

  // Log level: error | warn | info | debug
  logLevel: process.env.LOG_LEVEL || 'info',
};

module.exports = config;
