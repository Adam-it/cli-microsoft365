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

const expirationDateTimeMaxDays = 180;
const maxExpirationDateTime: Date = new Date();
maxExpirationDateTime.setDate(maxExpirationDateTime.getDate() + expirationDateTimeMaxDays);

const baseOptions = globalOptionsZod.extend({
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
  notificationUrl: zod.alias('n', z.string().optional()),
  expirationDateTime: zod.alias('e', z.string().optional()),
  clientState: zod.alias('c', z.string().optional())
});

const options = baseOptions;

type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class SpoListWebhookSetCommand extends SpoCommand {
  public get name(): string {
    return commands.LIST_WEBHOOK_SET;
  }

  public get description(): string {
    return 'Updates the specified webhook';
  }

  public get schema(): z.ZodTypeAny | undefined {
    return options;
  }

  public getRefinedSchema(schema: typeof options): z.ZodEffects<any> | undefined {
    return schema
      .strict()
      .refine(opts => [opts.listId, opts.listTitle, opts.listUrl].filter(option => option !== undefined).length === 1, {
        message: 'Specify exactly one of listId, listTitle or listUrl.',
        path: ['listId']
      })
      .refine(opts => opts.notificationUrl !== undefined || opts.expirationDateTime !== undefined || opts.clientState !== undefined, {
        message: 'Specify notificationUrl, expirationDateTime, clientState or multiple, at least one is required.',
        path: ['notificationUrl']
      })
      .refine(opts => {
        if (!opts.expirationDateTime) {
          return true;
        }

        const parsedDateTime = Date.parse(opts.expirationDateTime);
        if (Number.isNaN(parsedDateTime)) {
          return false;
        }

        const expirationDate = new Date(parsedDateTime);
        return expirationDate > new Date() && expirationDate < maxExpirationDateTime;
      }, {
        message: `Provide an expiration date which is a date time in the future and within 6 months from now. If specifying a date, use one of the following formats:\n  'YYYY-MM-DD'\n  'YYYY-MM-DDThh:mm'\n  'YYYY-MM-DDThh:mmZ'\n  'YYYY-MM-DDThh:mm±hh:mm'`,
        path: ['expirationDateTime']
      });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    if (this.verbose) {
      await logger.logToStderr(`Updating webhook ${args.options.id} belonging to list ${args.options.listId || args.options.listTitle || args.options.listUrl} located at site ${args.options.webUrl}...`);
    }

    let requestUrl: string = `${args.options.webUrl}/_api/web`;

    if (args.options.listId) {
      requestUrl += `/lists(guid'${formatting.encodeQueryParameter(args.options.listId)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listTitle) {
      requestUrl += `/lists/GetByTitle('${formatting.encodeQueryParameter(args.options.listTitle)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }
    else if (args.options.listUrl) {
      const listServerRelativeUrl: string = urlUtil.getServerRelativePath(args.options.webUrl, args.options.listUrl);
      requestUrl += `/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/Subscriptions('${formatting.encodeQueryParameter(args.options.id)}')`;
    }

    const requestBody: Record<string, string | undefined> = {};

    if (args.options.notificationUrl) {
      requestBody.notificationUrl = args.options.notificationUrl;
    }

    if (args.options.expirationDateTime) {
      requestBody.expirationDateTime = new Date(args.options.expirationDateTime).toISOString();
    }

    if (args.options.clientState) {
      requestBody.clientState = args.options.clientState;
    }

    const requestOptions: CliRequestOptions = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      data: requestBody,
      responseType: 'json'
    };

    try {
      await request.patch(requestOptions);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoListWebhookSetCommand();
