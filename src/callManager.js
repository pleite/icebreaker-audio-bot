'use strict';

/**
 * callManager.js — not used in the current chat-only implementation.
 *
 * This file is retained as a placeholder.  A voice/audio participant mode
 * (using Azure Communication Services Call Automation) could be wired in
 * here in the future by setting AUDIO_MODE=true and providing
 * ACS_CONNECTION_STRING + COGNITIVE_SERVICES_ENDPOINT.
 *
 * For the current simple chat agent, all icebreaker logic lives in:
 *   src/bot.js            — turn management and Teams chat interaction
 *   src/aiManager.js      — AI-generated responses
 *   src/icebreakerActivity.js — session state machine
 */

module.exports = {};
