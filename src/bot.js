'use strict';

/**
 * IcebreakerBot — a simple Teams chat agent.
 *
 * The bot conducts icebreaker activities entirely through the Teams meeting
 * chat.  No audio, no extra Azure services — just a Bot Framework webhook
 * and (optionally) an OpenAI key.
 *
 * Flow
 * ----
 * 1. Bot is added to a Teams meeting → auto-starts the icebreaker.
 * 2. Bot tries to discover the meeting roster via TeamsInfo; if that fails,
 *    participants are added dynamically as they type their first response.
 * 3. AI generates all text (intro, per-participant prompts, acknowledgements,
 *    wrap-up).  Falls back to scripted text when no AI key is configured.
 * 4. The bot keeps state in memory — one IcebreakerSession per conversation.
 *
 * Commands (type in the meeting chat)
 * ------------------------------------
 *   start   — start (or restart) the icebreaker
 *   skip    — skip the current participant's turn
 *   stop    — end the current session
 *   status  — show progress
 *   help    — show this list
 */

const { ActivityHandler, MessageFactory, TeamsInfo } = require('botbuilder');
const aiManager = require('./aiManager');
const { HINT } = require('./aiManager');
const { IcebreakerSession } = require('./icebreakerActivity');
const config = require('./config');
const logger = require('./logger');

class IcebreakerBot extends ActivityHandler {
  constructor() {
    super();

    /** @type {Map<string, IcebreakerSession>} conversationId → session */
    this._sessions = new Map();

    // Bot was added to a conversation / meeting chat
    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id === context.activity.recipient.id) {
          await this._onBotAdded(context);
        }
      }
      await next();
    });

    // Incoming message from a participant
    this.onMessage(async (context, next) => {
      await this._onMessage(context);
      await next();
    });
  }

  // ---------------------------------------------------------------------------
  // Bot joined
  // ---------------------------------------------------------------------------

  async _onBotAdded(context) {
    if (this._isInMeeting(context)) {
      // Auto-start when added directly to a meeting
      await this._startIcebreaker(context);
    } else {
      await context.sendActivity(
        MessageFactory.text(
          "👋 Hi! I'm the **Icebreaker Bot** 🎉\n\n" +
            'Invite me to a **Teams meeting** and I\'ll lead a fun icebreaker ' +
            'activity right in the chat — no setup required.\n\n' +
            'Type `help` for more info.'
        )
      );
    }
  }

  // ---------------------------------------------------------------------------
  // Message routing
  // ---------------------------------------------------------------------------

  async _onMessage(context) {
    // Teams wraps @mentions as <at>Name</at>. Strip those specifically.
    // Do NOT use a generic HTML-stripping regex — it can be bypassed and the
    // text value is never rendered as HTML anyway.
    const raw = (context.activity.text || '').replace(/<at>[^<]*<\/at>/g, '').trim();
    const lower = raw.toLowerCase();
    const convId = context.activity.conversation.id;
    const session = this._sessions.get(convId);

    // Commands — always handled regardless of session state
    if (lower === 'start' || lower === 'icebreaker' || lower === 'begin') {
      await this._startIcebreaker(context);
      return;
    }
    if (lower === 'help' || lower === '?') {
      await this._sendHelp(context);
      return;
    }
    if (lower === 'status') {
      await this._sendStatus(context, session);
      return;
    }
    if (lower === 'skip' || lower === 'next') {
      await this._skipCurrent(context, session);
      return;
    }
    if (lower === 'stop' || lower === 'end' || lower === 'cancel') {
      await this._stop(context, session);
      return;
    }

    // No active session — nudge the user
    if (!session || session.phase === 'ended') {
      await context.sendActivity(
        MessageFactory.text(
          "I'm the Icebreaker Bot! 🎉 Type `start` to kick off an icebreaker, or `help` for more info."
        )
      );
      return;
    }

    // Active session — treat the message as an icebreaker response
    await this._handleResponse(context, session, raw);
  }

  // ---------------------------------------------------------------------------
  // Start a new icebreaker session
  // ---------------------------------------------------------------------------

  async _startIcebreaker(context) {
    const convId = context.activity.conversation.id;
    const botId = context.activity.recipient.id;

    const session = new IcebreakerSession(convId, config.defaultActivity);
    this._sessions.set(convId, session);

    // Try to pre-populate the participant list from the meeting roster.
    // This requires the RSC permission "ChatMember.Read" in the manifest.
    // If it fails we fall back to dynamic discovery (participants join by typing).
    try {
      const { members } = await TeamsInfo.getPagedMembers(context);
      for (const m of members) {
        if (m.id !== botId) {
          session.addParticipant(m.id, m.name || m.givenName || 'Participant');
        }
      }
      logger.info(`Session ${convId}: roster loaded with ${session.participants.length} participants`);
    } catch (err) {
      logger.warn(
        `Session ${convId}: could not load roster (${err.message}) — ` +
          'participants will be added dynamically as they respond'
      );
    }

    session.phase = 'intro';

    const intro = await aiManager.generateResponse(session, HINT.INTRO);
    await context.sendActivity(MessageFactory.text(intro));

    if (session.participants.length === 0) {
      // Nobody in the roster yet — ask people to join
      const nudge = await aiManager.generateResponse(session, HINT.NO_PARTICIPANTS);
      await context.sendActivity(MessageFactory.text(nudge));
    }

    session.phase = 'activity';
    await this._promptNext(context, session);
  }

  // ---------------------------------------------------------------------------
  // Handle an incoming message as an icebreaker response
  // ---------------------------------------------------------------------------

  async _handleResponse(context, session, text) {
    if (session.phase !== 'activity') return;

    const senderId = context.activity.from.id;
    const senderName = context.activity.from.name || 'Participant';
    const botId = context.activity.recipient.id;

    // Ignore echoes of the bot's own messages
    if (senderId === botId) return;

    // Dynamically register this person if they aren't in the roster yet
    session.addParticipant(senderId, senderName);

    // Only accept from someone who still has a turn pending
    const participant = session.participants.find((p) => p.id === senderId && !p.hasGone);
    if (!participant) return;

    // Give the AI full context of what was said
    session.recordParticipantSpeech(senderName, text);
    session.markParticipantDone(senderId);
    session._waitingForId = null;

    logger.info(`Session ${session.sessionId}: response from ${senderName}`);

    const ack = await aiManager.generateResponse(session, HINT.ACKNOWLEDGEMENT, {
      participant,
      spokenText: text,
    });
    await context.sendActivity(MessageFactory.text(ack));

    await this._promptNext(context, session);
  }

  // ---------------------------------------------------------------------------
  // Prompt the next participant (or wrap up)
  // ---------------------------------------------------------------------------

  async _promptNext(context, session) {
    if (session.isActivityComplete()) {
      session.phase = 'wrapup';
      const wrapup = await aiManager.generateResponse(session, HINT.WRAPUP);
      await context.sendActivity(MessageFactory.text(wrapup));
      session.phase = 'ended';
      logger.info(`Session ${session.sessionId}: icebreaker complete`);
      return;
    }

    const next = session.getNextParticipant();
    if (!next) {
      // No roster yet — waiting for someone to respond
      return;
    }

    session._waitingForId = next.id;
    const prompt = await aiManager.generateResponse(session, HINT.PARTICIPANT_PROMPT, {
      participant: next,
    });
    await context.sendActivity(MessageFactory.text(prompt));
  }

  // ---------------------------------------------------------------------------
  // Commands
  // ---------------------------------------------------------------------------

  async _skipCurrent(context, session) {
    if (!session || session.phase !== 'activity') {
      await context.sendActivity(MessageFactory.text('No active icebreaker to skip.'));
      return;
    }
    const next = session.getNextParticipant();
    if (!next) {
      await context.sendActivity(MessageFactory.text('No participants left to skip!'));
      return;
    }

    session.markParticipantDone(next.id);
    session._waitingForId = null;

    const msg = await aiManager.generateResponse(session, HINT.SKIP, { participant: next });
    await context.sendActivity(MessageFactory.text(msg));
    await this._promptNext(context, session);
  }

  async _stop(context, session) {
    if (!session || session.phase === 'ended') {
      await context.sendActivity(MessageFactory.text('No active icebreaker session to stop.'));
      return;
    }
    session.phase = 'ended';
    this._sessions.delete(context.activity.conversation.id);
    await context.sendActivity(
      MessageFactory.text('Icebreaker session stopped. Type `start` to begin a new one! 👋')
    );
  }

  async _sendHelp(context) {
    await context.sendActivity(
      MessageFactory.text(
        '**Icebreaker Bot — Help** 🎉\n\n' +
          '**How it works:**\n' +
          '• Invite me to a Teams meeting and the icebreaker starts automatically\n' +
          '• Or type `start` in any meeting chat\n' +
          '• I go around and ask each person a question — just type your reply!\n' +
          '• If your roster isn\'t detected, everyone joins by typing their first response\n\n' +
          '**Commands:**\n' +
          '• `start` — begin (or restart) the icebreaker\n' +
          '• `skip` — skip the current participant\'s turn\n' +
          '• `stop` — end the current session\n' +
          '• `status` — show progress\n' +
          '• `help` — this message\n\n' +
          '**Activities** (chosen randomly each session):\n' +
          '• 🎭 Two Truths and a Lie\n' +
          '• 🤔 Would You Rather\n' +
          '• 💡 Fun Fact Round\n' +
          '• ⚡ Speed Introductions\n\n' +
          '_Customise the bot\'s personality by editing `instructions.txt`._'
      )
    );
  }

  async _sendStatus(context, session) {
    if (!session || session.phase === 'ended') {
      await context.sendActivity(
        MessageFactory.text('No active icebreaker session. Type `start` to begin!')
      );
      return;
    }
    const done = session.participants.filter((p) => p.hasGone).length;
    const total = session.participants.length;
    const next = session.getNextParticipant();
    const aiStatus = aiManager.isConfigured ? '✅ enabled' : '⚠️ fallback mode (no AI key)';
    await context.sendActivity(
      MessageFactory.text(
        '**Icebreaker Status**\n' +
          `• Activity: **${session.activityName}**\n` +
          `• Phase: **${session.phase}**\n` +
          `• Progress: **${done} / ${total || '?'}** participants done\n` +
          `• Waiting for: **${next?.name || 'anyone to respond'}**\n` +
          `• AI: ${aiStatus}`
      )
    );
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  _isInMeeting(context) {
    const cd = context.activity.channelData;
    return !!(cd?.meeting?.id || cd?.meeting?.joinUrl || cd?.onlineMeetingInfo);
  }
}

module.exports = { IcebreakerBot };
