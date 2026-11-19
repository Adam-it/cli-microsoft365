export function unavailable(name: string): (...args: any[]) => never {
  return (): never => {
    throw new Error(`'${name}' is not available in the browser build`);
  };
}
