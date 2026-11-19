// replaces chalk, which would add terminal color codes to the output.
// Supports the same chained calls, eg. chalk.red.bold('text'), and returns
// the text as is.
const chalk: any = new Proxy((...text: unknown[]): string => text.join(' '), {
  get: (): any => chalk
});

export default chalk;
