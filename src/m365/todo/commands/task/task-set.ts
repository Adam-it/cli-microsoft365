import { Logger } from '../../../../cli/Logger.js';
import GlobalOptions from '../../../../GlobalOptions.js';
import request, { CliRequestOptions } from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { validation } from '../../../../utils/validation.js';
import DelegatedGraphCommand from '../../../base/DelegatedGraphCommand.js';
import commands from '../../commands.js';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  listName?: string;
  listId?: string;
  title?: string;
  status?: string;
  bodyContent?: string;
  bodyContentType?: string;
  dueDateTime?: string;
  importance?: string;
  reminderDateTime?: string;
  categories?: string;
  completedDateTime?: string;
  startDateTime?: string;
}

class TodoTaskSetCommand extends DelegatedGraphCommand {
  public get name(): string {
    return commands.TASK_SET;
  }

  public get description(): string {
    return 'Update a task in a Microsoft To Do task list';
  }

  constructor() {
    super();

    this.#initTelemetry();
    this.#initOptions();
    this.#initValidators();
    this.#initOptionSets();
  }

  #initTelemetry(): void {
    this.telemetry.push((args: CommandArgs) => {
      Object.assign(this.telemetryProperties, {
        listId: typeof args.options.listId !== 'undefined',
        listName: typeof args.options.listName !== 'undefined',
        status: typeof args.options.status !== 'undefined',
        title: typeof args.options.title !== 'undefined',
        bodyContent: typeof args.options.bodyContent !== 'undefined',
        bodyContentType: args.options.bodyContentType,
        dueDateTime: typeof args.options.dueDateTime !== 'undefined',
        importance: args.options.importance,
        reminderDateTime: typeof args.options.reminderDateTime !== 'undefined',
        categories: typeof args.options.categories !== 'undefined',
        completedDateTime: typeof args.options.completedDateTime !== 'undefined',
        startDateTime: typeof args.options.startDateTime !== 'undefined'
      });
    });
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '-i, --id <id>'
      },
      {
        option: '-t, --title [title]'
      },
      {
        option: '-s, --status [status]',
        autocomplete: ['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']
      },
      {
        option: '--listName [listName]'
      },
      {
        option: '--listId [listId]'
      },
      {
        option: '--bodyContent [bodyContent]'
      },
      {
        option: '--bodyContentType [bodyContentType]',
        autocomplete: ['text', 'html']
      },
      {
        option: '--dueDateTime [dueDateTime]'
      },
      {
        option: '--importance [importance]',
        autocomplete: ['low', 'normal', 'high']
      },
      {
        option: '--reminderDateTime [reminderDateTime]'
      },
      {
        option: '--categories [categories]'
      },
      {
        option: '--completedDateTime [completedDateTime]'
      },
      {
        option: '--startDateTime [startDateTime]'
      }
    );
  }

  #initValidators(): void {
    this.validators.push(
      async (args: CommandArgs) => {
        if (args.options.status &&
          args.options.status !== 'notStarted' &&
          args.options.status !== 'inProgress' &&
          args.options.status !== 'completed' &&
          args.options.status !== 'waitingOnOthers' &&
          args.options.status !== 'deferred') {
          return `${args.options.status} is not a valid value. Allowed values are notStarted|inProgress|completed|waitingOnOthers|deferred`;
        }

        if (args.options.bodyContentType && ['text', 'html'].indexOf(args.options.bodyContentType.toLowerCase()) === -1) {
          return `'${args.options.bodyContentType}' is not a valid value for the bodyContentType option. Allowed values are text|html`;
        }

        if (args.options.importance && ['low', 'normal', 'high'].indexOf(args.options.importance.toLowerCase()) === -1) {
          return `'${args.options.importance}' is not a valid value for the importance option. Allowed values are low|normal|high`;
        }

        if (args.options.dueDateTime && !validation.isValidISODateTime(args.options.dueDateTime)) {
          return `'${args.options.dueDateTime}' is not a valid ISO date string`;
        }

        if (args.options.reminderDateTime && !validation.isValidISODateTime(args.options.reminderDateTime)) {
          return `'${args.options.reminderDateTime}' is not a valid ISO date string`;
        }

        if (args.options.completedDateTime && !validation.isValidISODateTime(args.options.completedDateTime)) {
          return `'${args.options.completedDateTime}' is not a valid datetime.`;
        }

        if (args.options.startDateTime && !validation.isValidISODateTime(args.options.startDateTime)) {
          return `'${args.options.startDateTime}' is not a valid datetime.`;
        }

        return true;
      }
    );
  }

  #initOptionSets(): void {
    this.optionSets.push({ options: ['listId', 'listName'] });
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    const endpoint: string = `${this.resource}/v1.0`;
    const data = this.mapRequestBody(args.options);

    try {
      const listId: string = await this.getTodoListId(args.options);
      const requestOptions: CliRequestOptions = {
        url: `${endpoint}/me/todo/lists/${listId}/tasks/${formatting.encodeQueryParameter(args.options.id)}`,
        headers: {
          accept: 'application/json;odata.metadata=none',
          'Content-Type': 'application/json'
        },
        data: data,
        responseType: 'json'
      };

      const res = await request.patch<any>(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }

  private async getTodoListId(options: GlobalOptions): Promise<string> {
    if (options.listId) {
      return options.listId;
    }

    const requestOptions: any = {
      url: `${this.resource}/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(options.listName!)}'`,
      headers: {
        accept: 'application/json;odata.metadata=none'
      },
      responseType: 'json'
    };

    const response = await request.get<{ value: [{ id: string }] }>(requestOptions);
    const taskList: { id: string } | undefined = response.value[0];

    if (!taskList) {
      throw `The specified task list does not exist`;
    }

    return taskList.id;
  }

  private getDateTimeTimeZone(dateTime: string): { dateTime: string, timeZone: string } {
    return {
      dateTime: dateTime,
      timeZone: 'Etc/GMT'
    };
  }

  private mapRequestBody(options: Options): any {
    const requestBody: any = {};

    if (options.status) {
      requestBody.status = options.status;
    }

    if (options.title) {
      requestBody.title = options.title;
    }

    if (options.importance) {
      requestBody.importance = options.importance.toLowerCase();
    }

    if (options.bodyContentType || options.bodyContent) {
      requestBody.body = {
        content: options.bodyContent,
        contentType: options.bodyContentType?.toLowerCase() || 'text'
      };
    }

    if (options.dueDateTime) {
      requestBody.dueDateTime = this.getDateTimeTimeZone(options.dueDateTime);
    }

    if (options.reminderDateTime) {
      requestBody.reminderDateTime = this.getDateTimeTimeZone(options.reminderDateTime);
    }

    if (options.categories) {
      requestBody.categories = options.categories.split(',');
    }

    if (options.completedDateTime) {
      requestBody.completedDateTime = this.getDateTimeTimeZone(options.completedDateTime);
    }

    if (options.startDateTime) {
      requestBody.startDateTime = this.getDateTimeTimeZone(options.startDateTime);
    }

    return requestBody;
  }
}

export default new TodoTaskSetCommand();