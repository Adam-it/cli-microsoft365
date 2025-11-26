import { z } from 'zod';
import { cli } from '../../../../cli/cli.js';
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
    }))),
    force: zod.alias('f', z.boolean().optional())
  })
  .strict();

const options = baseOptions;

type Options = z.infer<typeof options>;

class SpoListWebhookRemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_REMOVE;
  }

  public get description(): string {
    return 'Removes the specified webhook from the list';
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
    const listIdentifier = args.options.listId ?? args.options.listTitle ?? args.options.listUrl;
    const removeWebhook = async (): Promise<void> => {
      if (this.verbose) {
        await logger.logToStderr(`Webhook ${args.options.id} is about to be removed from list ${listIdentifier} located at site ${args.options.webUrl}...`);
      }

      const requestUrl = this.getRequestUrl(args.options);

      const requestOptions: CliRequestOptions = {
        url: requestUrl,
        method: 'DELETE',
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };

      try {
        await request.delete(requestOptions);
        // REST delete call doesn't return anything
      }
      catch (err: any) {
        this.handleRejectedODataJsonPromise(err);
      }
    };

    if (args.options.force) {
      await removeWebhook();
    }
    else {
      const result = await cli.promptForConfirmation({ message: `Are you sure you want to remove webhook ${args.options.id} from list ${listIdentifier} located at site ${args.options.webUrl}?` });

      if (result) {
        await removeWebhook();
      }
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

export default new SpoListWebhookRemoveCommand();
