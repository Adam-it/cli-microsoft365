// POSIX-style subset of the Node.js path module

export const sep = '/';
export const delimiter = ':';

export function isAbsolute(path: string): boolean {
  return path.startsWith('/');
}

export function normalize(path: string): string {
  const absolute = isAbsolute(path);
  const segments: string[] = [];

  path.split('/').forEach(segment => {
    if (segment === '' || segment === '.') {
      return;
    }

    if (segment === '..') {
      if (segments.length > 0 && segments[segments.length - 1] !== '..') {
        segments.pop();
      }
      else if (!absolute) {
        segments.push('..');
      }
      return;
    }

    segments.push(segment);
  });

  const normalized = segments.join('/');
  return absolute ? `/${normalized}` : normalized || '.';
}

export function join(...paths: string[]): string {
  return normalize(paths.filter(p => p).join('/'));
}

export function resolve(...paths: string[]): string {
  let resolved = '';

  for (let i = paths.length - 1; i >= 0 && !isAbsolute(resolved); i--) {
    if (paths[i]) {
      resolved = resolved ? `${paths[i]}/${resolved}` : paths[i];
    }
  }

  return normalize(isAbsolute(resolved) ? resolved : `/${resolved}`);
}

export function dirname(path: string): string {
  const trimmed = path.replace(/\/+$/, '');
  const pos = trimmed.lastIndexOf('/');
  if (pos < 0) {
    return '.';
  }

  return pos === 0 ? '/' : trimmed.substring(0, pos);
}

export function basename(path: string, suffix?: string): string {
  const trimmed = path.replace(/\/+$/, '');
  const base = trimmed.substring(trimmed.lastIndexOf('/') + 1);

  return suffix && base.endsWith(suffix) && base !== suffix
    ? base.substring(0, base.length - suffix.length)
    : base;
}

export function extname(path: string): string {
  const base = basename(path);
  const pos = base.lastIndexOf('.');

  return pos > 0 ? base.substring(pos) : '';
}

export function relative(from: string, to: string): string {
  const fromSegments = resolve(from).split('/').filter(s => s);
  const toSegments = resolve(to).split('/').filter(s => s);

  let common = 0;
  while (common < fromSegments.length &&
    common < toSegments.length &&
    fromSegments[common] === toSegments[common]) {
    common++;
  }

  return [
    ...fromSegments.slice(common).map(() => '..'),
    ...toSegments.slice(common)
  ].join('/');
}

const path = { sep, delimiter, isAbsolute, normalize, join, resolve, dirname, basename, extname, relative };

export const posix = path;
export default { ...path, posix };
