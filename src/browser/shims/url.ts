const _URL = globalThis.URL;
const _URLSearchParams = globalThis.URLSearchParams;

export function fileURLToPath(url: string | URL): string {
  return String(url);
}

export function pathToFileURL(path: string): URL {
  return new _URL(`file:///${path.replace(/^\/+/, '')}`);
}

export { _URL as URL, _URLSearchParams as URLSearchParams };
export default { fileURLToPath, pathToFileURL, URL: _URL, URLSearchParams: _URLSearchParams };
