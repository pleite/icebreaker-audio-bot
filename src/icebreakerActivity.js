'use strict';

/**
 * Icebreaker Activity Module
 *
 * Pure business logic for the icebreaker — no Azure/Teams dependencies.
 * All game state and conversation history live here so they can be unit tested
 * independently of any cloud infrastructure.
 *
 * The fixed script methods (getIntroScript, getParticipantPrompt, …) are kept
 * as fallbacks.  In normal operation, AiManager generates all spoken text and
 * records it here via addToHistory() so the AI model always has full context.
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const ACTIVITIES = Object.freeze({
  TWO_TRUTHS: 'two_truths_and_a_lie',
  WOULD_YOU_RATHER: 'would_you_rather',
  FUN_FACT: 'fun_fact_round',
  SPEED_INTRO: 'speed_intro',
});

const ACTIVITY_NAMES = Object.freeze({
  [ACTIVITIES.TWO_TRUTHS]: 'Two Truths and a Lie',
  [ACTIVITIES.WOULD_YOU_RATHER]: 'Would You Rather',
  [ACTIVITIES.FUN_FACT]: 'Fun Fact Round',
  [ACTIVITIES.SPEED_INTRO]: 'Speed Introductions',
});

const WOULD_YOU_RATHER_QUESTIONS = Object.freeze([
  'Would you rather be able to fly or be completely invisible?',
  'Would you rather always speak in rhymes or always sing instead of talking?',
  'Would you rather have a rewind button or a pause button for your life?',
  'Would you rather know the history of every object you touched or be able to talk to animals?',
  'Would you rather always be slightly overdressed or always be slightly underdressed?',
  'Would you rather be able to time travel or teleport anywhere instantly?',
  'Would you rather work from home forever or always work from an office?',
  'Would you rather have an extra hour every day or an extra day every week?',
  'Would you rather be famous but poor or unknown but wealthy?',
  'Would you rather master every musical instrument or speak every language fluently?',
  'Would you rather never have to sleep or never have to eat?',
  'Would you rather have a personal chef or a personal driver?',
  'Would you rather be able to read minds or predict the future?',
  'Would you rather have no internet for a month or no social media for a year?',
  'Would you rather live in the past or the future?',
]);

const ACKNOWLEDGEMENTS = Object.freeze([
  'Great answer!',
  'Interesting! I love it.',
  'That is fantastic!',
  'Wonderful, thank you for sharing!',
  'I did not see that coming!',
  'Very creative!',
  'Thank you for sharing that!',
  'Love it! Great response.',
  'Brilliant answer!',
  'That is so cool!',
]);

// ---------------------------------------------------------------------------
// Helper functions
// ---------------------------------------------------------------------------

/**
 * Returns a random element from an array.
 * @param {readonly any[]} arr
 * @param {number[]} [exclude=[]] - indices to exclude
 * @returns {{ value: any, index: number } | null}
 */
function randomFrom(arr, exclude = []) {
  const candidates = arr
    .map((value, index) => ({ value, index }))
    .filter(({ index }) => !exclude.includes(index));
  if (candidates.length === 0) return null;
  return candidates[Math.floor(Math.random() * candidates.length)];
}

/** Returns a random acknowledgement phrase. */
function getRandomAcknowledgement() {
  return randomFrom(ACKNOWLEDGEMENTS).value;
}

/** Returns a random activity type. */
function getRandomActivity() {
  const keys = Object.values(ACTIVITIES);
  return keys[Math.floor(Math.random() * keys.length)];
}

// ---------------------------------------------------------------------------
// IcebreakerSession — state machine for a single call
// ---------------------------------------------------------------------------

/**
 * Represents one active icebreaker session.
 * Used for both chat mode (keyed by conversation ID) and
 * audio mode (keyed by ACS call connection ID).
 *
 * Phases:
 *   connecting → intro → activity → wrapup → ended
 */
class IcebreakerSession {
  /**
   * @param {string} sessionId  - Conversation ID (chat) or ACS call connection ID (audio)
   * @param {string|null} [activity] - Activity type (default: random)
   */
  constructor(sessionId, activity = null) {
    /** Alias kept for backward-compat with callManager which reads callConnectionId */
    this.callConnectionId = sessionId;
    this.sessionId = sessionId;
    this.activity = activity && Object.values(ACTIVITIES).includes(activity)
      ? activity
      : getRandomActivity();
    /** @type {Array<{id: string, name: string, hasGone: boolean}>} */
    this.participants = [];
    this.phase = 'connecting';
    /** Track which "Would You Rather" question indices have been used */
    this._usedQuestionIndices = [];
    /** In chat mode: the participant ID the bot most recently prompted (for @mention UX) */
    this._waitingForId = null;
    /**
     * Conversation history for the AI model.
     * Format: [{ role: 'assistant'|'user', content: string }]
     * 'assistant' entries = what the bot has said.
     * 'user'      entries = what participants have said (prefixed with name).
     */
    this._conversationHistory = [];
  }

  // -------------------------------------------------------------------------
  // Conversation history (used by AiManager for context-aware generation)
  // -------------------------------------------------------------------------

  /**
   * Append an entry to the conversation history.
   * @param {'assistant'|'user'} role
   * @param {string} content
   */
  addToHistory(role, content) {
    this._conversationHistory.push({ role, content });
  }

  /**
   * Record a participant's spoken response.  Adds a 'user' history entry so
   * the AI knows what has already been said and can vary acknowledgements.
   * @param {string} participantName
   * @param {string} speech - Transcribed text from STT
   */
  recordParticipantSpeech(participantName, speech) {
    if (speech) {
      this._conversationHistory.push({
        role: 'user',
        content: `${participantName}: ${speech}`,
      });
    }
  }

  /**
   * Return a copy of the conversation history (safe to pass to OpenAI).
   * Capped at the most recent 20 messages to stay within token limits.
   * @returns {Array<{role: string, content: string}>}
   */
  getConversationHistory() {
    return this._conversationHistory.slice(-20);
  }

  // -------------------------------------------------------------------------
  // Computed properties
  // -------------------------------------------------------------------------

  /** Human-readable name of the current activity (used in AI prompts). */
  get activityName() {
    return ACTIVITY_NAMES[this.activity] || this.activity;
  }

  /** A random scripted acknowledgement phrase (used as AI fallback). */
  getRandomAcknowledgement() {
    return getRandomAcknowledgement();
  }

  // -------------------------------------------------------------------------
  // Participant management
  // -------------------------------------------------------------------------

  /** Add a participant if not already tracked. */
  addParticipant(id, name) {
    if (!id) return;
    if (!this.participants.find((p) => p.id === id)) {
      this.participants.push({ id, name: name || 'Participant', hasGone: false });
    }
  }

  /** Remove a participant (e.g., they left the call). */
  removeParticipant(id) {
    this.participants = this.participants.filter((p) => p.id !== id);
  }

  /** Return the first participant who hasn't had their turn yet. */
  getNextParticipant() {
    return this.participants.find((p) => !p.hasGone) || null;
  }

  /** Mark a participant as done and advance the queue. */
  markParticipantDone(id) {
    const participant = this.participants.find((p) => p.id === id);
    if (participant) participant.hasGone = true;
  }

  /** Returns true when every tracked participant has had a turn. */
  isActivityComplete() {
    return this.participants.length > 0 && this.participants.every((p) => p.hasGone);
  }

  // -------------------------------------------------------------------------
  // Script generation — all text the bot speaks out loud
  // -------------------------------------------------------------------------

  /** The opening announcement for the chosen activity. */
  getIntroScript() {
    const name = ACTIVITY_NAMES[this.activity];
    switch (this.activity) {
      case ACTIVITIES.TWO_TRUTHS:
        return (
          `Welcome everyone! I am your icebreaker host today. ` +
          `We are going to play ${name}! ` +
          `Here is how it works: each person will share three statements about themselves — ` +
          `two that are true and one that is a lie. ` +
          `Everyone else tries to guess which one is the lie. ` +
          `Let us start! I will go around and ask each of you in turn.`
        );
      case ACTIVITIES.WOULD_YOU_RATHER:
        return (
          `Welcome everyone! I am your icebreaker host. ` +
          `We are going to play ${name}! ` +
          `I will ask each of you a fun hypothetical question and you choose which option you prefer. ` +
          `There are no wrong answers — just go with your gut! ` +
          `Let us begin.`
        );
      case ACTIVITIES.FUN_FACT:
        return (
          `Welcome everyone! I am your icebreaker host. ` +
          `Let us do a ${name} round! ` +
          `Each person will share one interesting or surprising fact about themselves. ` +
          `It can be a hobby, a hidden talent, an unusual experience — anything that surprises us! ` +
          `Ready? Let us get started.`
        );
      case ACTIVITIES.SPEED_INTRO:
        return (
          `Welcome everyone! I am your icebreaker host. ` +
          `We are doing Speed Introductions! ` +
          `Each person has about thirty seconds to share their name, where they are from, ` +
          `and one thing they are passionate about. Short and sweet! ` +
          `Let us go.`
        );
      default:
        return `Welcome everyone! I am your icebreaker host. Let us get to know each other!`;
    }
  }

  /**
   * Returns the prompt the bot speaks to a specific participant.
   * @param {{ id: string, name: string }} participant
   * @returns {string}
   */
  getParticipantPrompt(participant) {
    const { name } = participant;
    switch (this.activity) {
      case ACTIVITIES.TWO_TRUTHS:
        return (
          `${name}, it is your turn! ` +
          `Please share two truths and one lie about yourself. ` +
          `Go ahead whenever you are ready.`
        );
      case ACTIVITIES.WOULD_YOU_RATHER: {
        const result = randomFrom(WOULD_YOU_RATHER_QUESTIONS, this._usedQuestionIndices);
        if (result) {
          this._usedQuestionIndices.push(result.index);
          return `${name}, here is your question: ${result.value} Which would you choose and why?`;
        }
        // All questions used — reset and pick any
        this._usedQuestionIndices = [];
        const fallback = randomFrom(WOULD_YOU_RATHER_QUESTIONS);
        return `${name}, here is your question: ${fallback.value} Which would you choose?`;
      }
      case ACTIVITIES.FUN_FACT:
        return `${name}, please share your fun fact with us!`;
      case ACTIVITIES.SPEED_INTRO:
        return (
          `${name}, you are up! ` +
          `Tell us your name, where you are from, and one thing you are passionate about. ` +
          `Go ahead!`
        );
      default:
        return `${name}, it is your turn!`;
    }
  }

  /** The closing statement the bot speaks after all participants have gone. */
  getWrapUpScript() {
    return (
      `That is everyone! ` +
      `Thank you all so much for participating in our icebreaker today. ` +
      `Now that you know a little more about each other, ` +
      `I hope this meeting goes wonderfully. ` +
      `Have a fantastic session! I will leave you to it.`
    );
  }

  /** Message spoken when a participant did not respond in time. */
  getSkipScript(participantName) {
    return `No worries, ${participantName}! We can come back to you later. Let us move on.`;
  }
}

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------

module.exports = {
  ACTIVITIES,
  ACTIVITY_NAMES,
  WOULD_YOU_RATHER_QUESTIONS,
  ACKNOWLEDGEMENTS,
  IcebreakerSession,
  getRandomActivity,
  getRandomAcknowledgement,
};
