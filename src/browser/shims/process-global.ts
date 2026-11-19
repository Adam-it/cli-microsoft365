// injected by the browser build so that references to the process global
// resolve to this object

const noop = (): void => {
  // nothing to do in a browser
};

const stream = {
  isTTY: false,
  columns: 80,
  write: (): boolean => true,
  on: noop,
  once: noop,
  off: noop,
  removeListener: noop
};

const hrtime: any = (): [number, number] => {
  const now = performance.now();
  return [Math.floor(now / 1e3), Math.round((now % 1e3) * 1e6)];
};
hrtime.bigint = (): bigint => BigInt(Math.round(performance.now() * 1e6));

const processShim: any = {
  env: {},
  argv: [],
  execArgv: [],
  platform: 'browser',
  version: '',
  versions: {},
  pid: 0,
  ppid: 0,
  exitCode: undefined,
  stdout: stream,
  stderr: stream,
  stdin: stream,
  hrtime,
  cwd: (): string => '/',
  uptime: (): number => performance.now() / 1e3,
  nextTick: (callback: (...args: any[]) => void, ...args: any[]): void => queueMicrotask(() => callback(...args)),
  emitWarning: noop,
  exit: (code?: number): never => {
    throw new Error(`process.exit(${code ?? 0}) was called, which is not possible in the browser build`);
  },
  on: (): any => processShim,
  once: (): any => processShim,
  off: (): any => processShim,
  removeListener: (): any => processShim
};

export { processShim as process };
