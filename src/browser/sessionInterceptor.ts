import Axios, { AxiosError, AxiosResponse, InternalAxiosRequestConfig } from 'axios';

export interface BrowserTransportRequest {
  url: string;
  method: string;
  headers: { [name: string]: string };
  body?: string;
}

export interface BrowserTransportResponse {
  status: number;
  statusText: string;
  /** Header names must be lower case */
  headers: { [name: string]: string };
  body: string;
}

/**
 * Sends an HTTP request on behalf of the CLI. A host passes one when requests
 * can't be sent from the page that runs the CLI, for example to send them
 * from a SharePoint tab instead of from an extension page.
 */
export type BrowserTransport = (request: BrowserTransportRequest) => Promise<BrowserTransportResponse>;

interface Session {
  hosts: Set<string>;
  transport?: BrowserTransport;
}

interface Digest {
  value: string;
  expires: number;
}

let session: Session | undefined;
const digests = new Map<string, Digest>();

export function setSession(hosts: string[], transport?: BrowserTransport): void {
  session = {
    hosts: new Set(hosts.map(h => h.toLowerCase())),
    transport
  };
  digests.clear();
}

export function isSessionHost(url: string): boolean {
  try {
    return session !== undefined && session.hosts.has(new URL(url).hostname.toLowerCase());
  }
  catch {
    return false;
  }
}

/**
 * Returns the URL of the web a SharePoint API URL points to: the part before
 * /_api/ or /_vti_bin/. Returns undefined for other URLs.
 */
export function getWebUrl(url: string): string | undefined {
  const path = url.split('?')[0];
  const match = /\/(_api|_vti_bin)(\/|$)/i.exec(path);
  return match ? path.substring(0, match.index) : undefined;
}

async function send(request: BrowserTransportRequest): Promise<BrowserTransportResponse> {
  if (session?.transport) {
    return session.transport(request);
  }

  const res = await fetch(request.url, {
    method: request.method,
    credentials: 'include',
    headers: request.headers,
    body: request.body
  });
  const headers: { [name: string]: string } = {};
  res.headers.forEach((value, name) => headers[name.toLowerCase()] = value);

  return {
    status: res.status,
    statusText: res.statusText,
    headers,
    body: await res.text()
  };
}

async function getDigest(webUrl: string): Promise<string> {
  const cached = digests.get(webUrl);
  if (cached && cached.expires > Date.now()) {
    return cached.value;
  }

  // not sent through the CLI's request object so that it doesn't come back
  // through this interceptor
  const res = await send({
    url: `${webUrl}/_api/contextinfo`,
    method: 'POST',
    headers: { accept: 'application/json;odata=nometadata' }
  });

  if (res.status < 200 || res.status > 299) {
    throw new Error(`Could not get a request digest for ${webUrl} (HTTP ${res.status}). Check that you're signed in to ${new URL(webUrl).origin} in this browser: open it in a tab, then try again.`);
  }

  const contextInfo: { FormDigestValue: string; FormDigestTimeoutSeconds: number } = JSON.parse(res.body);
  digests.set(webUrl, {
    value: contextInfo.FormDigestValue,
    expires: Date.now() + (contextInfo.FormDigestTimeoutSeconds - 30) * 1000
  });

  return contextInfo.FormDigestValue;
}

async function transportAdapter(config: InternalAxiosRequestConfig): Promise<AxiosResponse> {
  const data: unknown = config.data;
  if (data !== undefined && data !== null && typeof data !== 'string') {
    throw new AxiosError('Only text request bodies can be sent through the host transport', AxiosError.ERR_BAD_REQUEST, config);
  }

  const headers: { [name: string]: string } = {};
  Object.entries(config.headers.toJSON()).forEach(([name, value]) => {
    if (value !== undefined && value !== null) {
      headers[name] = String(value);
    }
  });

  const res = await send({
    url: Axios.getUri(config),
    method: (config.method ?? 'get').toUpperCase(),
    headers,
    body: data ?? undefined
  });

  const response: AxiosResponse = {
    data: res.body,
    status: res.status,
    statusText: res.statusText,
    headers: res.headers,
    config
  };

  if (!config.validateStatus || config.validateStatus(res.status)) {
    return response;
  }

  throw new AxiosError(
    `Request failed with status code ${res.status}`,
    res.status >= 500 ? AxiosError.ERR_BAD_RESPONSE : AxiosError.ERR_BAD_REQUEST,
    config,
    undefined,
    response
  );
}

/**
 * Makes requests to the SharePoint hosts of the current session work with the
 * session cookies of the signed-in user: sends credentials, removes headers
 * that browsers don't allow and adds a request digest to writes. The CLI
 * itself only sends a digest on CSOM calls, because with a bearer token REST
 * writes don't need one.
 */
export async function sessionInterceptor(config: InternalAxiosRequestConfig): Promise<InternalAxiosRequestConfig> {
  const url: string = config.url ?? '';
  if (!session || !isSessionHost(url)) {
    return config;
  }

  config.withCredentials = true;
  config.headers.delete('user-agent');
  config.headers.delete('accept-encoding');

  if (session.transport) {
    config.adapter = transportAdapter;
  }

  const method: string = (config.method ?? 'get').toLowerCase();
  if (method === 'get' ||
    method === 'head' ||
    config.headers.has('X-RequestDigest') ||
    /\/_api\/contextinfo$/i.test(url.split('?')[0])) {
    return config;
  }

  const webUrl: string | undefined = getWebUrl(url);
  if (webUrl) {
    config.headers.set('X-RequestDigest', await getDigest(webUrl));
  }

  return config;
}
