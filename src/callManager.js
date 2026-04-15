'use strict';

/**
 * CallManager — Azure Communication Services Call Automation integration.
 *
 * Responsibilities:
 *  - Join Teams meetings as an audio participant via ACS
 *  - Play TTS prompts using Azure Neural voices
 *  - Trigger speech recognition to hear participant responses
 *  - Drive the icebreaker state machine in response to ACS callback events
 */

const { CallAutomationClient } = require('@azure/communication-call-automation');
const config = require('./config');
const logger = require('./logger');
const { IcebreakerSession, getRandomAcknowledgement } = require('./icebreakerActivity');

// ---------------------------------------------------------------------------
// ACS event type strings
// ---------------------------------------------------------------------------

const EVENT = Object.freeze({
  CALL_CONNECTED: 'Microsoft.Communication.CallConnected',
  CALL_DISCONNECTED: 'Microsoft.Communication.CallDisconnected',
  PARTICIPANTS_UPDATED: 'Microsoft.Communication.ParticipantsUpdated',
  PLAY_COMPLETED: 'Microsoft.Communication.PlayCompleted',
  PLAY_FAILED: 'Microsoft.Communication.PlayFailed',
  RECOGNIZE_COMPLETED: 'Microsoft.Communication.RecognizeCompleted',
  RECOGNIZE_FAILED: 'Microsoft.Communication.RecognizeFailed',
  RECOGNIZE_CANCELED: 'Microsoft.Communication.RecognizeCanceled',
});

// Operation context tags — embedded in each ACS media request so we know
// which step of the flow just finished when the callback fires.
const CTX = Object.freeze({
  INTRO: 'intro_complete',
  ACKNOWLEDGEMENT: 'acknowledgement_complete',
  SKIP: 'skip_complete',
  WRAPUP: 'wrapup_complete',
});

class CallManager {
  constructor() {
    /** @type {CallAutomationClient} */
    this._client = new CallAutomationClient(config.acsConnectionString);

    /** @type {Map<string, IcebreakerSession>} callConnectionId → session */
    this._sessions = new Map();
  }

  // -------------------------------------------------------------------------
  // Public: join a Teams meeting
  // -------------------------------------------------------------------------

  /**
   * Join a Teams meeting as an ACS bot participant.
   *
   * @param {string} meetingUrl - The Teams meeting join URL (from the bot activity)
   * @param {string|null} [activity] - Optional activity type override
   * @returns {Promise<string>} The ACS call connection ID
   */
  async joinTeamsMeeting(meetingUrl, activity = null) {
    logger.info(`Joining Teams meeting: ${meetingUrl}`);

    const callbackUrl = `${config.callbackBaseUrl}/api/callbacks`;

    const result = await this._client.joinCall(
      // Locator: tell ACS this is a Teams meeting link
      { kind: 'teamsMeetingLinkLocator', meetingLink: meetingUrl },
      callbackUrl,
      {
        // Link our Cognitive Services resource so ACS can do TTS + STT
        callIntelligenceOptions: {
          cognitiveServicesEndpoint: config.cognitiveServicesEndpoint,
        },
      }
    );

    const callConnectionId = result.callConnectionProperties.callConnectionId;
    const session = new IcebreakerSession(callConnectionId, activity || config.defaultActivity);
    this._sessions.set(callConnectionId, session);

    logger.info(`Joined call. ID: ${callConnectionId}, Activity: ${session.activity}`);
    return callConnectionId;
  }

  // -------------------------------------------------------------------------
  // Public: process an incoming ACS callback event (called from HTTP handler)
  // -------------------------------------------------------------------------

  /**
   * Dispatch an ACS event to the appropriate handler.
   *
   * @param {object} event - The raw ACS CloudEvent payload
   */
  async handleEvent(event) {
    const type = event.type || event.eventType;
    const data = event.data || {};
    const callConnectionId = data.callConnectionId;

    logger.debug(`ACS event [${type}] call=${callConnectionId}`);

    switch (type) {
      case EVENT.CALL_CONNECTED:
        await this._onCallConnected(callConnectionId);
        break;

      case EVENT.CALL_DISCONNECTED:
        await this._onCallDisconnected(callConnectionId);
        break;

      case EVENT.PARTICIPANTS_UPDATED:
        await this._onParticipantsUpdated(callConnectionId, data.participants || []);
        break;

      case EVENT.PLAY_COMPLETED:
        await this._onPlayCompleted(callConnectionId, data.operationContext);
        break;

      case EVENT.PLAY_FAILED:
        logger.warn(`Play failed for call ${callConnectionId}: ${data.resultInformation?.message}`);
        // Treat as completed so the flow continues
        await this._onPlayCompleted(callConnectionId, data.operationContext);
        break;

      case EVENT.RECOGNIZE_COMPLETED: {
        const speech = data.recognizeResult?.speechResult?.speech || '';
        await this._onRecognizeCompleted(callConnectionId, speech, data.operationContext);
        break;
      }

      case EVENT.RECOGNIZE_FAILED:
      case EVENT.RECOGNIZE_CANCELED:
        await this._onRecognizeFailed(callConnectionId, data.operationContext);
        break;

      default:
        logger.debug(`Unhandled ACS event type: ${type}`);
    }
  }

  // -------------------------------------------------------------------------
  // Internal event handlers
  // -------------------------------------------------------------------------

  async _onCallConnected(callConnectionId) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    session.phase = 'intro';
    logger.info(`Call connected [${callConnectionId}], playing intro`);
    await this._playTts(callConnectionId, session.getIntroScript(), CTX.INTRO);
  }

  async _onCallDisconnected(callConnectionId) {
    this._sessions.delete(callConnectionId);
    logger.info(`Call disconnected [${callConnectionId}]`);
  }

  async _onParticipantsUpdated(callConnectionId, participants) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    for (const p of participants) {
      // Skip the bot itself (communicationUser kind = the ACS identity we joined as)
      if (!p.identifier || p.identifier.kind === 'communicationUser') continue;

      const id =
        p.identifier.microsoftTeamsUserId ||
        p.identifier.phoneNumber ||
        p.identifier.rawId;

      if (id) {
        const name = p.displayName || 'Participant';
        session.addParticipant(id, name);
        logger.info(`Participant tracked [${callConnectionId}]: ${name} (${id})`);
      }
    }
  }

  async _onPlayCompleted(callConnectionId, operationContext) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    switch (operationContext) {
      case CTX.INTRO:
        session.phase = 'activity';
        await this._promptNextParticipant(callConnectionId);
        break;

      case CTX.ACKNOWLEDGEMENT:
      case CTX.SKIP:
        await this._promptNextParticipant(callConnectionId);
        break;

      case CTX.WRAPUP:
        await this._hangUp(callConnectionId);
        break;

      default:
        logger.debug(`play completed, unrecognised context: ${operationContext}`);
    }
  }

  async _onRecognizeCompleted(callConnectionId, speech, operationContext) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    logger.info(`Recognized speech [${callConnectionId}]: "${speech}"`);

    // Extract participant ID from context (format: "recognize_<participantId>")
    const participantId = operationContext?.replace(/^recognize_/, '');
    if (participantId) session.markParticipantDone(participantId);

    await this._playTts(callConnectionId, getRandomAcknowledgement(), CTX.ACKNOWLEDGEMENT);
  }

  async _onRecognizeFailed(callConnectionId, operationContext) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    const participantId = operationContext?.replace(/^recognize_/, '');
    const participant = session.participants.find((p) => p.id === participantId);
    const name = participant?.name || 'participant';

    logger.warn(`Recognition failed [${callConnectionId}] for ${name}`);

    if (participantId) session.markParticipantDone(participantId);

    await this._playTts(callConnectionId, session.getSkipScript(name), CTX.SKIP);
  }

  // -------------------------------------------------------------------------
  // Internal helpers
  // -------------------------------------------------------------------------

  /**
   * Find the next participant who hasn't had a turn and start recognizing
   * their speech, or play the wrap-up message if everyone has gone.
   */
  async _promptNextParticipant(callConnectionId) {
    const session = this._sessions.get(callConnectionId);
    if (!session) return;

    if (session.isActivityComplete()) {
      session.phase = 'wrapup';
      await this._playTts(callConnectionId, session.getWrapUpScript(), CTX.WRAPUP);
      return;
    }

    const next = session.getNextParticipant();
    if (!next) {
      // No tracked participants yet — play a nudge and wait for people to join
      await this._playTts(
        callConnectionId,
        "I don't see anyone in the participant list yet. " +
          'Please make sure I have permission to see meeting participants, ' +
          'and feel free to say hello to get started!',
        CTX.WRAPUP
      );
      return;
    }

    const prompt = session.getParticipantPrompt(next);
    logger.info(`Prompting participant [${callConnectionId}]: ${next.name}`);

    try {
      const callMedia = this._client.getCallConnection(callConnectionId).getCallMedia();

      await callMedia.startRecognizing(
        // Target the specific participant
        this._buildParticipantIdentifier(next.id),
        /* maxTonesToCollect */ 1,
        {
          recognizeInputType: 'speech',
          playPrompt: {
            kind: 'textSource',
            text: prompt,
            voiceName: config.botVoice,
          },
          endSilenceTimeoutInMs: config.recognizeSilenceTimeoutMs,
          initialSilenceTimeoutInMs: config.recognizeInitialSilenceTimeoutMs,
          locale: 'en-US',
          operationContext: `recognize_${next.id}`,
        }
      );
    } catch (err) {
      logger.error(`startRecognizing failed for ${next.name}: ${err.message}`);
      session.markParticipantDone(next.id);
      // Skip and try the next person
      await this._promptNextParticipant(callConnectionId);
    }
  }

  /**
   * Play a TTS message to all participants in the call.
   */
  async _playTts(callConnectionId, text, operationContext = '') {
    logger.debug(`TTS [${callConnectionId}] ctx=${operationContext}: "${text.slice(0, 60)}..."`);
    try {
      const callMedia = this._client.getCallConnection(callConnectionId).getCallMedia();

      await callMedia.playToAll(
        [{ kind: 'textSource', text, voiceName: config.botVoice }],
        { loop: false, operationContext }
      );
    } catch (err) {
      logger.error(`playToAll failed [${callConnectionId}]: ${err.message}`);
    }
  }

  /**
   * Gracefully disconnect the bot from the call.
   */
  async _hangUp(callConnectionId) {
    logger.info(`Hanging up [${callConnectionId}]`);
    try {
      // false = only disconnect this bot, not the entire meeting
      await this._client.getCallConnection(callConnectionId).hangUp(false);
    } catch (err) {
      logger.warn(`hangUp failed [${callConnectionId}]: ${err.message}`);
    } finally {
      this._sessions.delete(callConnectionId);
    }
  }

  /**
   * Build the ACS CommunicationIdentifier for a Teams user.
   */
  _buildParticipantIdentifier(participantId) {
    // Teams users are identified by their AAD Object ID
    if (participantId.startsWith('+')) {
      return { kind: 'phoneNumber', phoneNumber: participantId };
    }
    return { kind: 'microsoftTeamsUser', microsoftTeamsUserId: participantId };
  }

  /** Expose the sessions map for diagnostics (read-only snapshot). */
  getActiveSessions() {
    return [...this._sessions.entries()].map(([id, s]) => ({
      callConnectionId: id,
      activity: s.activity,
      phase: s.phase,
      participantCount: s.participants.length,
    }));
  }
}

// Export a singleton — one CallManager per process is sufficient
module.exports = new CallManager();
