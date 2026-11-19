/**
 * Splits a command line into arguments. Single- and double-quoted spans are
 * kept as one argument with the quotes removed, so JSON option values such as
 * --fields '{"Title":"x"}' survive intact. Inside double quotes \" stands for
 * a literal double quote.
 */
export function tokenize(line: string): string[] {
  const args: string[] = [];
  let current: string = '';
  let inArg: boolean = false;
  let quote: string | undefined;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (quote) {
      if (char === quote) {
        quote = undefined;
      }
      else if (char === '\\' && quote === '"' && line[i + 1] === '"') {
        current += '"';
        i++;
      }
      else {
        current += char;
      }
      continue;
    }

    if (char === '"' || char === "'") {
      quote = char;
      inArg = true;
      continue;
    }

    if (/\s/.test(char)) {
      if (inArg) {
        args.push(current);
        current = '';
        inArg = false;
      }
      continue;
    }

    current += char;
    inArg = true;
  }

  if (quote) {
    throw new Error(`Unterminated ${quote} quote in the command line`);
  }

  if (inArg) {
    args.push(current);
  }

  return args;
}
