export function setTimeout<T = void>(delay?: number, value?: T): Promise<T> {
  return new Promise<T>(resolve => globalThis.setTimeout(() => resolve(value as T), delay));
}

export default { setTimeout };
