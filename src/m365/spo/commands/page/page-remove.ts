import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import request, { CliRequestOptions } from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { urlUtil } from '../../../../utils/urlUtil';
import { validation } from '../../../../utils/validation';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
  webUrl: string;
  recycle?: boolean;
  bypassSharedLock?: boolean;
  force?: boolean;
}

class SpoPageRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.PAGE_REMOVE;
  }

  public get description(): string {
    return 'Removes a modern page';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initTypes();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        force: !!args.options.force,
        recycle: !!args.options.recycle,
        bypassSharedLock: !!args.options.bypassSharedLock
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-u, --webUrl <webUrl>'
      },
      {
        option: '-n, --name <name>'
      },
      {
        option: '--recycle'
      },
      {
        option: '--bypassSharedLock'
      },
      {
        option: '-f, --force'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => validation.isValidSharePointUrl(args.options.webUrl)
    );
  }

  #initTypes(): void {
    this.types.string.push('name', 'webUrl');
    this.types.boolean.push('force', 'bypassSharedLock', 'recycle');
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (args.options.force) {
      await this.removePage(logger, args);
    }
    else {
      const result = await Cli.prompt<{ continue: boolean }>(
        {
          type: 'confirm',
          name: 'continue',
          default: false,
          message: `Are you sure you want to remove the page '${args.options.name}'?`
        });

      if (result.continue) {
        await this.removePage(logger, args);
      }
    }
  }

  private async removePage(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      let pageName: string = args.options.name;

      if (!pageName.toLowerCase().endsWith('.aspx')) {
        pageName += '.aspx';
      }

      if (this.verbose) {
        logger.logToStderr(`Removing page ${pageName}...`);
      }

      const filePath = `${urlUtil.getServerRelativeSiteUrl(args.options.webUrl)}/SitePages/${pageName}`;
      const requestOptions: CliRequestOptions = {
        url: `${args.options.webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl='${formatting.encodeQueryParameter(filePath)}')`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      if (args.options.bypassSharedLock) {
        requestOptions.headers!.Prefer = 'bypass-shared-lock';
      }
      if (args.options.recycle) {
        requestOptions.url += '/Recycle';

        await request.post(requestOptions);
      }
      else {
        await request.delete(requestOptions);
      }
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoPageRemoveCommand();
