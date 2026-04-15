'use strict';

const {
  ACTIVITIES,
  ACTIVITY_NAMES,
  IcebreakerSession,
  getRandomActivity,
  getRandomAcknowledgement,
} = require('../src/icebreakerActivity');

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

describe('ACTIVITIES', () => {
  test('contains four activity types', () => {
    expect(Object.keys(ACTIVITIES)).toHaveLength(4);
  });

  test('every activity has a human-readable name', () => {
    for (const key of Object.values(ACTIVITIES)) {
      expect(ACTIVITY_NAMES[key]).toBeTruthy();
    }
  });
});

// ---------------------------------------------------------------------------
// Module-level helpers
// ---------------------------------------------------------------------------

describe('getRandomActivity', () => {
  test('returns a valid activity key', () => {
    const result = getRandomActivity();
    expect(Object.values(ACTIVITIES)).toContain(result);
  });

  test('returns varied results over many calls', () => {
    const results = new Set(Array.from({ length: 50 }, () => getRandomActivity()));
    expect(results.size).toBeGreaterThan(1);
  });
});

describe('getRandomAcknowledgement', () => {
  test('returns a non-empty string', () => {
    expect(typeof getRandomAcknowledgement()).toBe('string');
    expect(getRandomAcknowledgement().length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// IcebreakerSession
// ---------------------------------------------------------------------------

describe('IcebreakerSession constructor', () => {
  test('assigns a random activity when none provided', () => {
    const session = new IcebreakerSession('conv-1');
    expect(Object.values(ACTIVITIES)).toContain(session.activity);
  });

  test('uses the specified activity when valid', () => {
    const session = new IcebreakerSession('conv-1', ACTIVITIES.FUN_FACT);
    expect(session.activity).toBe(ACTIVITIES.FUN_FACT);
  });

  test('falls back to random activity when an invalid value is provided', () => {
    const session = new IcebreakerSession('conv-1', 'not_a_real_activity');
    expect(Object.values(ACTIVITIES)).toContain(session.activity);
  });

  test('starts in connecting phase', () => {
    const session = new IcebreakerSession('conv-1');
    expect(session.phase).toBe('connecting');
  });

  test('sessionId and callConnectionId are both set', () => {
    const session = new IcebreakerSession('conv-abc');
    expect(session.sessionId).toBe('conv-abc');
    expect(session.callConnectionId).toBe('conv-abc');
  });
});

// ---------------------------------------------------------------------------
// activityName getter
// ---------------------------------------------------------------------------

describe('IcebreakerSession.activityName', () => {
  test('returns the human-readable name', () => {
    const session = new IcebreakerSession('c1', ACTIVITIES.TWO_TRUTHS);
    expect(session.activityName).toBe('Two Truths and a Lie');
  });
});

// ---------------------------------------------------------------------------
// Participant management
// ---------------------------------------------------------------------------

describe('IcebreakerSession participant management', () => {
  let session;
  beforeEach(() => {
    session = new IcebreakerSession('conv-1', ACTIVITIES.FUN_FACT);
  });

  test('addParticipant adds a new participant', () => {
    session.addParticipant('u1', 'Alice');
    expect(session.participants).toHaveLength(1);
    expect(session.participants[0]).toMatchObject({ id: 'u1', name: 'Alice', hasGone: false });
  });

  test('addParticipant is idempotent — same ID not added twice', () => {
    session.addParticipant('u1', 'Alice');
    session.addParticipant('u1', 'Alice again');
    expect(session.participants).toHaveLength(1);
  });

  test('addParticipant ignores falsy IDs', () => {
    session.addParticipant(null, 'Ghost');
    session.addParticipant('', 'Empty');
    expect(session.participants).toHaveLength(0);
  });

  test('addParticipant defaults name to "Participant" when missing', () => {
    session.addParticipant('u1', '');
    expect(session.participants[0].name).toBe('Participant');
  });

  test('removeParticipant removes the correct entry', () => {
    session.addParticipant('u1', 'Alice');
    session.addParticipant('u2', 'Bob');
    session.removeParticipant('u1');
    expect(session.participants).toHaveLength(1);
    expect(session.participants[0].id).toBe('u2');
  });

  test('getNextParticipant returns the first who has not gone', () => {
    session.addParticipant('u1', 'Alice');
    session.addParticipant('u2', 'Bob');
    expect(session.getNextParticipant()).toMatchObject({ id: 'u1' });
  });

  test('getNextParticipant returns null when everyone is done', () => {
    session.addParticipant('u1', 'Alice');
    session.markParticipantDone('u1');
    expect(session.getNextParticipant()).toBeNull();
  });

  test('markParticipantDone sets hasGone to true', () => {
    session.addParticipant('u1', 'Alice');
    session.markParticipantDone('u1');
    expect(session.participants[0].hasGone).toBe(true);
  });

  test('isActivityComplete is false when no one has gone', () => {
    session.addParticipant('u1', 'Alice');
    expect(session.isActivityComplete()).toBe(false);
  });

  test('isActivityComplete is false when participants list is empty', () => {
    expect(session.isActivityComplete()).toBe(false);
  });

  test('isActivityComplete is true when all participants are done', () => {
    session.addParticipant('u1', 'Alice');
    session.addParticipant('u2', 'Bob');
    session.markParticipantDone('u1');
    session.markParticipantDone('u2');
    expect(session.isActivityComplete()).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Conversation history
// ---------------------------------------------------------------------------

describe('IcebreakerSession conversation history', () => {
  let session;
  beforeEach(() => {
    session = new IcebreakerSession('conv-1');
  });

  test('addToHistory stores entries', () => {
    session.addToHistory('assistant', 'Hello everyone!');
    session.addToHistory('user', 'Alice: I love hiking.');
    expect(session.getConversationHistory()).toHaveLength(2);
  });

  test('getConversationHistory returns correct roles and content', () => {
    session.addToHistory('assistant', 'Hello!');
    const history = session.getConversationHistory();
    expect(history[0]).toEqual({ role: 'assistant', content: 'Hello!' });
  });

  test('getConversationHistory caps at 20 most recent entries', () => {
    for (let i = 0; i < 25; i++) {
      session.addToHistory('assistant', `Message ${i}`);
    }
    const history = session.getConversationHistory();
    expect(history).toHaveLength(20);
    expect(history[19].content).toBe('Message 24');
  });

  test('recordParticipantSpeech adds a user entry with name prefix', () => {
    session.recordParticipantSpeech('Alice', 'I enjoy photography.');
    const history = session.getConversationHistory();
    expect(history[0]).toEqual({ role: 'user', content: 'Alice: I enjoy photography.' });
  });

  test('recordParticipantSpeech ignores empty speech', () => {
    session.recordParticipantSpeech('Alice', '');
    expect(session.getConversationHistory()).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// Fallback scripts (used when AI is not configured)
// ---------------------------------------------------------------------------

describe('IcebreakerSession fallback scripts', () => {
  test('getIntroScript returns a non-empty string for each activity', () => {
    for (const activity of Object.values(ACTIVITIES)) {
      const session = new IcebreakerSession('c1', activity);
      expect(typeof session.getIntroScript()).toBe('string');
      expect(session.getIntroScript().length).toBeGreaterThan(10);
    }
  });

  test('getParticipantPrompt includes the participant name', () => {
    const session = new IcebreakerSession('c1', ACTIVITIES.FUN_FACT);
    const prompt = session.getParticipantPrompt({ id: 'u1', name: 'Alice' });
    expect(prompt).toContain('Alice');
  });

  test('getWrapUpScript returns a non-empty string', () => {
    const session = new IcebreakerSession('c1');
    expect(session.getWrapUpScript().length).toBeGreaterThan(10);
  });

  test('getSkipScript includes the participant name', () => {
    const session = new IcebreakerSession('c1');
    expect(session.getSkipScript('Bob')).toContain('Bob');
  });

  test('getRandomAcknowledgement (instance method) returns a non-empty string', () => {
    const session = new IcebreakerSession('c1');
    expect(session.getRandomAcknowledgement().length).toBeGreaterThan(0);
  });
});
