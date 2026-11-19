function assert(value: unknown, message?: string | Error): asserts value {
  if (!value) {
    throw message instanceof Error ? message : new Error(message ?? 'Assertion failed');
  }
}

export function equal(actual: unknown, expected: unknown, message?: string | Error): void {
  // eslint-disable-next-line eqeqeq
  assert(actual == expected, message);
}

export function strictEqual(actual: unknown, expected: unknown, message?: string | Error): void {
  assert(actual === expected, message);
}

export const ok: (value: unknown, message?: string | Error) => asserts value = assert;

assert.ok = ok;
assert.equal = equal;
assert.strictEqual = strictEqual;

export default assert;
