import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import { globalOptionsZod } from '../../../../Command.js';
import { formatting } from '../../../../utils/formatting.js';
import { odata } from '../../../../utils/odata.js';
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
    listId: zod.alias('i', z.string().optional().refine(id => id === undefined || validation.isValidGuid(id), id => ({
      message: `'${id}' is not a valid GUID.`
    }))),
    listTitle: zod.alias('t', z.string().optional()),
    listUrl: z.string().optional()
  })
  .strict();

const options = baseOptions;

type Options = z.infer<typeof options>;

class SpoListWebhookListCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_LIST;
  }

  public get description(): string {
    return 'Lists all webhooks for the specified list';
  }


  public defaultProperties(): string[] | undefined {
    return ['id', 'clientState', 'expirationDateTime', 'resource'];
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
      await logger.logToStderr(`Retrieving webhook information for list ${args.options.listTitle || args.options.listId || args.options.listUrl} in site at ${args.options.webUrl}...`);
    }

    const requestUrl: string = this.getRequestUrl(args.options);

    try {
      const webhooks = await odata.getAllItems<{ id: string; clientState?: string; expirationDateTime: Date; resource: string }>(requestUrl);

      webhooks.forEach(webhook => {
        webhook.clientState = webhook.clientState ?? '';
      });

      await logger.log(webhooks);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private getRequestUrl(options: Options): string {
    let requestUrl = `${options.webUrl}/_api/web`;

    if (options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(options.listId)}')/Subscriptions`;
    }
    else if (options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(options.listTitle)}')/Subscriptions`;
    }
    else if (options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(options.webUrl, options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions`;
    }

    return requestUrl;
  }
}

export default new SpoListWebhookListCommand();
