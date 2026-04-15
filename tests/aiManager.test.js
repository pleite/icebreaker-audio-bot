'use strict';

/**
 * Tests for AiManager — uses a mock OpenAI client so no real API calls are made.
 */

const { AiManager, HINT } = require('../src/aiManager');
const { IcebreakerSession, ACTIVITIES } = require('../src/icebreakerActivity');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function makeSession(activity = ACTIVITIES.FUN_FACT, participants = []) {
  const session = new IcebreakerSession('test-conv', activity);
  session.phase = 'activity';
  for (const { id, name } of participants) {
    session.addParticipant(id, name);
  }
  return session;
}

function makeMockClient(responseText = 'Great question!') {
  return {
    chat: {
      completions: {
        create: jest.fn().mockResolvedValue({
          choices: [{ message: { content: responseText } }],
        }),
      },
    },
  };
}

// ---------------------------------------------------------------------------
// HINT constants
// ---------------------------------------------------------------------------

describe('HINT constants', () => {
  test('exports the expected keys', () => {
    expect(HINT.INTRO).toBe('intro');
    expect(HINT.PARTICIPANT_PROMPT).toBe('participant_prompt');
    expect(HINT.ACKNOWLEDGEMENT).toBe('acknowledgement');
    expect(HINT.SKIP).toBe('skip');
    expect(HINT.WRAPUP).toBe('wrapup');
    expect(HINT.NO_PARTICIPANTS).toBe('no_participants');
  });
});

// ---------------------------------------------------------------------------
// AiManager.isConfigured
// ---------------------------------------------------------------------------

describe('AiManager.isConfigured', () => {
  test('is false when no AI env vars are set', () => {
    const mgr = new AiManager();
    // By default in test env, no AI keys are set
    expect(mgr.isConfigured).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// generateResponse — with mocked AI client
// ---------------------------------------------------------------------------

describe('AiManager.generateResponse', () => {
  function makeManager(responseText) {
    const mgr = new AiManager();
    mgr._client = makeMockClient(responseText);
    mgr._model = 'gpt-test';
    return mgr;
  }

  test('returns AI-generated text', async () => {
    const mgr = makeManager('Welcome to the icebreaker!');
    const session = makeSession();
    session.phase = 'connecting';
    const result = await mgr.generateResponse(session, HINT.INTRO);
    expect(result).toBe('Welcome to the icebreaker!');
  });

  test('adds AI response to conversation history', async () => {
    const mgr = makeManager('Great fun fact!');
    const session = makeSession(ACTIVITIES.FUN_FACT, [{ id: 'u1', name: 'Alice' }]);
    await mgr.generateResponse(session, HINT.ACKNOWLEDGEMENT, {
      participant: { id: 'u1', name: 'Alice' },
      spokenText: 'I can juggle.',
    });
    const history = session.getConversationHistory();
    expect(history).toHaveLength(1);
    expect(history[0]).toEqual({ role: 'assistant', content: 'Great fun fact!' });
  });

  test('sends conversation history to the AI', async () => {
    const mgr = makeManager('Alice, your turn!');
    const session = makeSession(ACTIVITIES.WOULD_YOU_RATHER, [{ id: 'u1', name: 'Alice' }]);
    session.addToHistory('assistant', 'Welcome everyone!');

    await mgr.generateResponse(session, HINT.PARTICIPANT_PROMPT, {
      participant: { id: 'u1', name: 'Alice' },
    });

    const createArgs = mgr._client.chat.completions.create.mock.calls[0][0];
    // messages array: [system, ...history(1), user]
    expect(createArgs.messages.length).toBeGreaterThanOrEqual(3);
    const historyMsg = createArgs.messages.find(
      (m) => m.role === 'assistant' && m.content === 'Welcome everyone!'
    );
    expect(historyMsg).toBeTruthy();
  });

  test('falls back to scripted text when AI returns empty string', async () => {
    const mgr = makeManager('');
    const session = makeSession();
    const result = await mgr.generateResponse(session, HINT.INTRO);
    // Should return non-empty fallback
    expect(result.length).toBeGreaterThan(0);
  });

  test('falls back to scripted text when AI call throws', async () => {
    const mgr = new AiManager();
    mgr._client = {
      chat: { completions: { create: jest.fn().mockRejectedValue(new Error('Network error')) } },
    };
    mgr._model = 'gpt-test';
    const session = makeSession();
    const result = await mgr.generateResponse(session, HINT.WRAPUP);
    expect(result.length).toBeGreaterThan(0);
  });

  test('uses scripted fallback when no client is configured', async () => {
    const mgr = new AiManager(); // no _client set
    const session = makeSession(ACTIVITIES.SPEED_INTRO, [{ id: 'u1', name: 'Bob' }]);
    const result = await mgr.generateResponse(session, HINT.PARTICIPANT_PROMPT, {
      participant: { id: 'u1', name: 'Bob' },
    });
    expect(result).toContain('Bob');
  });
});

// ---------------------------------------------------------------------------
// Instructions loading
// ---------------------------------------------------------------------------

describe('AiManager instructions loading', () => {
  const originalEnv = process.env;

  afterEach(() => {
    process.env = { ...originalEnv };
  });

  test('uses ICEBREAKER_INSTRUCTIONS env var when set', () => {
    process.env = { ...originalEnv, ICEBREAKER_INSTRUCTIONS: 'You are a pirate host.' };
    const mgr = new AiManager();
    expect(mgr._instructions).toBe('You are a pirate host.');
  });

  test('falls back to default instructions when env var is absent', () => {
    const env = { ...originalEnv };
    delete env.ICEBREAKER_INSTRUCTIONS;
    delete env.ICEBREAKER_INSTRUCTIONS_FILE;
    process.env = env;
    const mgr = new AiManager();
    expect(mgr._instructions.length).toBeGreaterThan(0);
  });
});
