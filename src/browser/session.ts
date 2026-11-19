import auth from '../Auth.js';
import { cli } from '../cli/cli.js';
import { CommandInfo } from '../cli/CommandInfo.js';
import Command, { CommandError } from '../Command.js';
import request from '../request.js';
import { settingsNames } from '../settingsNames.js';
import { ConfirmationConfig, prompt } from '../utils/prompt.js';
import { BrowserTransport, sessionInterceptor, setSession } from './sessionInterceptor.js';

export interface BrowserCommandInfo {
  name: string;
  description: string;
  file: string;
  help?: string;
  aliases?: string[];
}

export type CommandLoader = () => Promise<{ default: unknown }>;

export interface BrowserHostOptions {
  /**
   * Root URL of the tenant's SharePoint host, eg. https://contoso.sharepoint.com
   */
  spoUrl: string;
  /**
   * Sends HTTP requests on behalf of the CLI. When not set, requests are sent
   * from the page that runs the CLI.
   */
  transport?: BrowserTransport;
  /**
   * Called when a command asks for confirmation. When not set, commands that
   * ask for confirmation fail unless they're run with --force.
   */
  confirm?: (message: string) => Promise<boolean>;
}

export const registry: {
  infos: BrowserCommandInfo[];
  loaders: { [file: string]: CommandLoader };
} = {
  infos: [],
  loaders: {}
};

let initialized: boolean = false;
let allowedHosts: string[] = [];
let confirmHandler: ((message: string) => Promise<boolean>) | undefined;

const settings = new Map<string, unknown>([
  [settingsNames.prompt, false],
  [settingsNames.output, 'json'],
  [settingsNames.disableTelemetry, true],
  [settingsNames.showHelpOnFailure, false],
  [settingsNames.autoOpenLinksInBrowser, false]
]);

// stands in for Configstore, which stores settings in a file
const config = {
  get: (key: string): unknown => settings.get(key),
  set: (key: string, value: unknown): void => { settings.set(key, value); },
  has: (key: string): boolean => settings.has(key),
  delete: (key: string): void => { settings.delete(key); },
  clear: (): void => settings.clear(),
  get all(): { [key: string]: unknown } { return Object.fromEntries(settings); }
};

function getSessionHosts(spoUrl: string): string[] {
  const hostname: string = new URL(spoUrl).hostname;
  const [tenant, ...rest] = hostname.split('.');
  const suffix: string = rest.join('.');

  return [hostname, `${tenant}-admin.${suffix}`, `${tenant}-my.${suffix}`];
}

function setupOnce(): void {
  // without a bearer token, the CLI sends requests without the Authorization
  // header. Requests to SharePoint are then authenticated with the session
  // cookies. For other resources there's nothing to authenticate with.
  auth.ensureAccessToken = async (resource: string): Promise<string> => {
    let hostname: string = '';
    try {
      hostname = new URL(resource).hostname.toLowerCase();
    }
    catch {
      // resource is not a URL
    }

    if (allowedHosts.includes(hostname)) {
      return '';
    }

    throw new Error(`This command requires access to ${resource}, which is not available in SharePoint session mode. Only the SharePoint APIs of the current tenant can be used.`);
  };
  // there's no connection information to restore or store
  const noop = async (): Promise<void> => {
    // nothing to do in a browser
  };
  auth.restoreAuth = noop;
  auth.storeConnectionInfo = noop;

  cli.getConfig = (): any => config;

  cli.loadAllCommandsInfo = (): void => {
    cli.commands = registry.infos as unknown as CommandInfo[];
  };
  cli.loadCommandFromFile = async (commandFile: string): Promise<void> => {
    const loader = registry.loaders[commandFile];
    if (!loader) {
      cli.currentCommandName = undefined;
      return;
    }

    const command = await loader();
    if (command.default instanceof Command) {
      const commandInfo = cli.commands.find(c => c.file === commandFile);
      cli.commandToExecute = cli.getCommandInfo(command.default, commandFile, commandInfo?.help);
    }
  };

  // in Node this ends with process.exit
  cli.closeWithError = async (error: any): Promise<void> => {
    throw error instanceof CommandError ? error : new CommandError(error?.message ?? String(error));
  };

  prompt.forConfirmation = async (config: ConfirmationConfig): Promise<boolean> => {
    if (confirmHandler) {
      return confirmHandler(config.message);
    }

    throw new Error(`This command asks for confirmation (${config.message}), which is not available here. Run it again with --force to confirm.`);
  };
  prompt.forInput = async (): Promise<string> => {
    throw new Error('Prompts are not available here. Specify all required options in the command line.');
  };
  prompt.forSelection = async <T>(): Promise<T> => {
    throw new Error('Prompts are not available here. Specify all required options in the command line.');
  };

  request.addRequestInterceptor(sessionInterceptor);
}

/**
 * Configures the CLI to run in a browser as the user who's signed in to
 * SharePoint in that browser. Can be called again to switch to another tenant.
 */
export function init(options: BrowserHostOptions): void {
  const spoUrl: string = new URL(options.spoUrl).origin;

  if (!initialized) {
    setupOnce();
    initialized = true;
  }

  allowedHosts = getSessionHosts(spoUrl).map(h => h.toLowerCase());
  confirmHandler = options.confirm;
  setSession(allowedHosts, options.transport);

  auth.connection.active = true;
  auth.connection.spoUrl = spoUrl;
  auth.connection.spoTenantId = undefined;
  auth.connection.accessTokens = {};
}

export function isInitialized(): boolean {
  return initialized;
}
