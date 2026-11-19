export function inspect(value: unknown): string {
  if (typeof value === 'string') {
    return value;
  }

  try {
    return JSON.stringify(value, null, 2);
  }
  catch {
    return String(value);
  }
}

export function promisify(fn: (...args: any[]) => void): (...args: any[]) => Promise<any> {
  return (...args: any[]): Promise<any> => new Promise((resolve, reject) => {
    fn(...args, (err: unknown, result: unknown) => err ? reject(err) : resolve(result));
  });
}

export function format(...args: unknown[]): string {
  return args.map(a => inspect(a)).join(' ');
}

export function inherits(ctor: any, superCtor: any): void {
  Object.setPrototypeOf(ctor.prototype, superCtor.prototype);
}

export function deprecate<T>(fn: T): T {
  return fn;
}

const _TextEncoder = globalThis.TextEncoder;
const _TextDecoder = globalThis.TextDecoder;

export { _TextEncoder as TextEncoder, _TextDecoder as TextDecoder };
export default { inspect, promisify, format, inherits, deprecate, TextEncoder: _TextEncoder, TextDecoder: _TextDecoder };
