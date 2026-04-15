'use strict';

/**
 * IcebreakerBot — Microsoft Bot Framework Teams bot.
 *
 * Handles:
 *  - Meeting lifecycle events (bot added to a meeting)
 *  - @mention messages (manual trigger, status, help)
 *  - Extracts the Teams meeting join URL and hands it to callManager
 */

const { ActivityHandler, MessageFactory, TurnContext } = require('botbuilder');
const callManager = require('./callManager');
const logger = require('./logger');
const { ACTIVITIES } = require('./icebreakerActivity');

class IcebreakerBot extends ActivityHandler {
  constructor() {
    super();

    // Fired when the bot (or any user) is added to a conversation
    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded || []) {
        if (member.id === context.activity.recipient.id) {
          // The bot itself was added
          await this._onBotJoined(context);
        }
      }
      await next();
    });

    // Fired when the bot receives a text message
    this.onMessage(async (context, next) => {
      await this._onMessage(context);
      await next();
    });
  }

  // -------------------------------------------------------------------------
  // Event handlers
  // -------------------------------------------------------------------------

  /** Bot was added to a conversation / meeting chat. */
  async _onBotJoined(context) {
    const meetingUrl = this._extractMeetingUrl(context);

    if (meetingUrl) {
      logger.info('Bot added to meeting — joining audio channel');
      await context.sendActivity(
        MessageFactory.text(
          '👋 Hi everyone! I\'m the **Icebreaker Bot** 🎉\n\n' +
            'I\'m joining the meeting audio right now. ' +
            'Once I\'m in, I\'ll lead a fun icebreaker activity to help everyone get to know each other!'
        )
      );
      await this._joinMeeting(context, meetingUrl);
    } else {
      await context.sendActivity(
        MessageFactory.text(
          '👋 Hi! I\'m the **Icebreaker Bot** 🎉\n\n' +
            'Add me to a **Teams meeting** and I\'ll join the audio to lead a fun icebreaker activity.\n\n' +
            'You can also type `start` in the meeting chat to kick things off manually.'
        )
      );
    }
  }

  /** Handle an incoming text message. */
  async _onMessage(context) {
    const text = (context.activity.text || '').toLowerCase().trim();
    const meetingUrl = this._extractMeetingUrl(context);

    if (text.includes('start') || text.includes('icebreaker') || text.includes('join')) {
      if (meetingUrl) {
        await context.sendActivity(
          MessageFactory.text('🎤 Starting the icebreaker! Joining the audio now… get ready! 🎉')
        );
        await this._joinMeeting(context, meetingUrl);
      } else {
        await context.sendActivity(
          MessageFactory.text(
            'Please use this command inside a Teams meeting chat so I can find the meeting URL. ' +
              'Alternatively, invite me to the meeting directly.'
          )
        );
      }
    } else if (text.includes('help')) {
      await context.sendActivity(
        MessageFactory.text(
          '**Icebreaker Bot — Help**\n\n' +
            '• Invite me to a Teams meeting and I will join automatically\n' +
            '• Type `start` in the meeting chat to trigger manually\n' +
            '• I will speak to each participant in turn and lead a fun icebreaker game\n\n' +
            '**Available activities** (chosen randomly each meeting):\n' +
            '• 🎭 Two Truths and a Lie\n' +
            '• 🤔 Would You Rather\n' +
            '• 💡 Fun Fact Round\n' +
            '• ⚡ Speed Introductions'
        )
      );
    } else if (text.includes('status')) {
      const sessions = callManager.getActiveSessions();
      if (sessions.length === 0) {
        await context.sendActivity(MessageFactory.text('No active icebreaker sessions right now.'));
      } else {
        const lines = sessions.map(
          (s) =>
            `• Call \`${s.callConnectionId.slice(0, 8)}…\` — ` +
            `activity: **${s.activity}**, phase: **${s.phase}**, ` +
            `participants: **${s.participantCount}**`
        );
        await context.sendActivity(
          MessageFactory.text(`**Active sessions:**\n${lines.join('\n')}`)
        );
      }
    } else {
      await context.sendActivity(
        MessageFactory.text(
          'I\'m the Icebreaker Bot! 🎉 Type `help` for usage info, or `start` to begin an icebreaker.'
        )
      );
    }
  }

  // -------------------------------------------------------------------------
  // Helpers
  // -------------------------------------------------------------------------

  /**
   * Extract the Teams meeting join URL from the activity's channelData.
   * Teams populates channelData.meeting.joinUrl when the conversation is
   * associated with an online meeting.
   *
   * @param {TurnContext} context
   * @returns {string|null}
   */
  _extractMeetingUrl(context) {
    const cd = context.activity.channelData;
    return (
      cd?.meeting?.joinUrl ||
      cd?.onlineMeetingInfo?.joinUrl ||
      null
    );
  }

  /**
   * Ask callManager to join the Teams meeting and handle errors gracefully.
   *
   * @param {TurnContext} context
   * @param {string} meetingUrl
   * @param {string|null} [activity]
   */
  async _joinMeeting(context, meetingUrl, activity = null) {
    try {
      const callConnectionId = await callManager.joinTeamsMeeting(meetingUrl, activity);
      logger.info(`Successfully joined meeting, callConnectionId=${callConnectionId}`);
    } catch (err) {
      logger.error(`Failed to join meeting: ${err.message}`);
      await context.sendActivity(
        MessageFactory.text(
          '⚠️ I had trouble joining the meeting audio. Please check the bot configuration and try again. ' +
            'Common issues: missing ACS_CONNECTION_STRING or COGNITIVE_SERVICES_ENDPOINT in environment.'
        )
      );
    }
  }
}

module.exports = { IcebreakerBot };
