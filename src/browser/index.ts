import yargs from 'yargs-parser';
import z from 'zod';
import { cli, CommandOutput } from '../cli/cli.js';
import { CommandInfo } from '../cli/CommandInfo.js';
import { settingsNames } from '../settingsNames.js';
import { BrowserCommandInfo, CommandLoader, init, isInitialized, registry } from './session.js';
import { tokenize } from './tokenize.js';

export type { BrowserCommandInfo, BrowserHostOptions, CommandLoader } from './session.js';
export type { BrowserTransport, BrowserTransportRequest, BrowserTransportResponse } from './sessionInterceptor.js';
export { init };

export interface SkippedCommand {
  name: string;
  reason: string;
}

export interface CommandListener {
  stdout?: (message: any) => void;
  stderr?: (message: any) => void;
}

export interface BrowserCommand {
  name: string;
  description: string;
  aliases?: string[];
  helpUrl?: string;
}

const docsUrl: string = 'https://pnp.github.io/cli-microsoft365/cmd/';
let skippedCommands: SkippedCommand[] = [];
let running: boolean = false;

/**
 * Registers the commands that are part of the bundle. Called by the generated
 * bundle entry point.
 * @param infos Information about the commands, as in allCommands.json
 * @param loaders Functions that load the command modules, by command file
 * @param skipped Commands left out of the bundle, with the reason
 */
export function registerCommands(infos: BrowserCommandInfo[], loaders: { [file: string]: CommandLoader }, skipped: SkippedCommand[] = []): void {
  registry.infos = infos;
  registry.loaders = loaders;
  skippedCommands = skipped;
}

function getHelpUrl(help: string | undefined): string | undefined {
  if (!help) {
    return undefined;
  }

  return docsUrl + help.replace(/\\/g, '/').replace(/\.[^/.]+$/, '');
}

export function listCommands(): BrowserCommand[] {
  return registry.infos.map(i => ({
    name: i.name,
    description: i.description,
    aliases: i.aliases,
    helpUrl: getHelpUrl(i.help)
  }));
}

export function listSkippedCommands(): SkippedCommand[] {
  return skippedCommands;
}

// same as in cli.ts, where it's not exported
function isMissingRequiredValueError(error: z.core.$ZodIssue, options: any): boolean {
  if (error.code !== 'invalid_type' && error.code !== 'invalid_value') {
    return false;
  }

  let currentValue: any = options;
  for (const pathSegment of error.path) {
    if (currentValue === null ||
      typeof currentValue !== 'object' ||
      !Object.prototype.hasOwnProperty.call(currentValue, pathSegment)) {
      return true;
    }

    currentValue = currentValue[pathSegment];
  }

  return false;
}

function getErrorMessage(error: any): string {
  if (typeof error === 'string') {
    return error;
  }

  if (error === undefined || error === null) {
    return 'Unknown error';
  }

  // CommandError wraps whatever the command threw in its message property
  if (typeof error.message === 'string') {
    return error.message;
  }

  if (error.message !== undefined && error.message !== null) {
    return getErrorMessage(error.message);
  }

  try {
    return JSON.stringify(error);
  }
  catch {
    return String(error);
  }
}

function getCommandNotFoundMessage(commandName: string): string {
  const skipped = skippedCommands.find(c => c.name === commandName);
  if (skipped) {
    return `Command '${commandName}' is not available in the browser: ${skipped.reason}.`;
  }

  if (!commandName.startsWith('spo')) {
    return `Command '${commandName}' was not found. Only SharePoint (spo) commands are available in the browser.`;
  }

  return `Command '${commandName}' was not found.`;
}

async function run(line: string, listener?: CommandListener): Promise<CommandOutput> {
  const rawArgs: string[] = tokenize(line);
  if (rawArgs[0] === 'm365' || rawArgs[0] === 'microsoft365') {
    rawArgs.shift();
  }

  let showHelp: boolean = false;
  if (rawArgs[0] === 'help') {
    showHelp = true;
    rawArgs.shift();
  }

  if (rawArgs.length === 0) {
    throw new Error('Specify a command to run, eg. spo web get --url https://contoso.sharepoint.com');
  }

  // loadCommandFromArgs doesn't reset these when it can't find a command
  cli.commandToExecute = undefined;
  cli.currentCommandName = undefined;
  cli.loadAllCommandsInfo();

  const parsedArgs: yargs.Arguments = yargs(rawArgs);
  await cli.loadCommandFromArgs(parsedArgs._);

  const commandInfo = cli.commandToExecute as CommandInfo | undefined;
  if (!commandInfo) {
    throw new Error(getCommandNotFoundMessage(parsedArgs._.join(' ')));
  }

  if (showHelp || parsedArgs.h || parsedArgs.help) {
    const help: string = `${commandInfo.name}: ${commandInfo.description}. Help is not available here, see ${getHelpUrl(commandInfo.help)}`;
    listener?.stdout?.(help);

    return { stdout: help, stderr: '' };
  }

  const options: yargs.Arguments = cli.getCommandOptionsFromArgs(rawArgs, commandInfo);
  delete (options as any)._;
  delete (options as any)['--'];
  cli.optionsFromArgs = { options };

  await commandInfo.command.processOptions(options);

  if (options.output === undefined) {
    options.output = cli.getSettingWithDefaultValue<string | undefined>(settingsNames.output, 'json');
  }

  let finalArgs: any = options;
  if (commandInfo.command.schema) {
    const result = await commandInfo.command.getSchemaToParse()!.safeParseAsync(options);
    if (!result.success) {
      throw new Error(result.error.issues
        .map(issue => {
          const message = isMissingRequiredValueError(issue, options)
            ? 'Required option not specified'
            : issue.message;

          return issue.path.length > 0
            ? `${issue.path.join('.')}: ${message}`
            : message;
        })
        .join('\n'));
    }

    finalArgs = result.data;
  }
  else {
    const validationResult = await commandInfo.command.validate({ options }, commandInfo);
    if (validationResult !== true) {
      throw new Error(String(validationResult).trim());
    }
  }

  return await cli.executeCommandWithOutput(commandInfo.command, { options: finalArgs }, listener);
}

/**
 * Runs a command specified the same way as in a terminal, eg.
 * spo web get --url https://contoso.sharepoint.com. Follows the
 * non-interactive path of cli.execute, but rejects with an Error instead of
 * printing the error and exiting the process.
 */
export async function executeCommandLine(line: string, listener?: CommandListener): Promise<CommandOutput> {
  if (!isInitialized()) {
    throw new Error('Call init before running commands');
  }

  if (running) {
    throw new Error('Another command is still running. Commands can only be run one at a time.');
  }

  running = true;
  // messages written outside of the command's logger
  const cliLog = cli.log;
  const cliError = cli.error;
  cli.log = (message?: any): void => listener?.stdout?.(message ?? '');
  cli.error = async (message?: any): Promise<void> => listener?.stderr?.(message ?? '');

  try {
    return await run(line, listener);
  }
  catch (e: any) {
    // executeCommandWithOutput rejects with { error, stderr }
    const error: Error & { stderr?: string } = new Error(getErrorMessage(e && e.error !== undefined ? e.error : e));
    if (e && typeof e.stderr === 'string') {
      error.stderr = e.stderr;
    }

    throw error;
  }
  finally {
    cli.log = cliLog;
    cli.error = cliError;
    running = false;
  }
}
