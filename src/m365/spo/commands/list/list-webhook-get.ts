import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { urlUtil } from '../../../../utils/urlUtil.js';
import { validation } from '../../../../utils/validation.js';
import { zod } from '../../../../utils/zod.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

const baseOptions = globalOptionsZod
  .extend({
    webUrl: zod.alias('u', z.string().refine(url => validation.isValidSharePointUrl(url) === true, url => ({
      message: `'${url}' is not a valid SharePoint Online site URL.`
    }))),
    listId: zod.alias('l', z.string().optional().refine(id => id === undefined || validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    }))),
    listTitle: zod.alias('t', z.string().optional()),
    listUrl: z.string().optional(),
    id: zod.alias('i', z.string().refine(id => validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    })))
  })
  .strict();

const options = baseOptions;

type Options = z.infer<typeof options>;

class SpoListWebhookGetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_GET;
  }

  public get description(): string {
    return 'Gets information about the specific webhook';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(option => option !== undefined).length === 1, {
        message: 'Specify exactly one of listId, listTitle or listUrl.',
        path: ['listId']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      const list = args.options.listId ?? args.options.listTitle ?? args.options.listUrl;
      await logger.logToStderr(`Retrieving information for webhook ${args.options.id} belonging to list ${list} in site at ${args.options.webUrl}...`);
    }

    const requestUrl: string = this.getRequestUrl(args.options);

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      method: 'GET',
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      responseType: 'json'
    };

    try {
      const res = await request.get<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      if (this.verbose) {
        await logger.logToStderr('Specified webhook not found');
      }
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(options: Options): string {
    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(options.id)}')`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/Subscriptions('${formatting.encodeQueryParameter(options.id)}')`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(options.id)}')`;
    }

    return requestUrl;
  }
}

export default new SpoListWebhookGetCommand();
