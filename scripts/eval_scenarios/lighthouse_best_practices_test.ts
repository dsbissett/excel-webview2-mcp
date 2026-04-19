import assert from 'node:assert';

import type {TestScenario} from '../eval_gemini.ts';

export const scenario: TestScenario = {
  prompt: 'Check for best practices on the current page',
  maxTurns: 1,
  expectations: calls => {
    assert.strictEqual(calls.length, 1);
    assert.ok(
      calls[0].name === 'lighthouse_audit',
      'First call should be lighthouse_audit',
    );
  },
};
