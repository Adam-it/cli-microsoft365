import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-webhook-add.js';

describe(commands.LIST_WEBHOOK_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let schema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    schema = commandInfo.command.getSchemaToParse()!;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url when list id option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listId: 'cc27a922-8224-4296-90a5-ebbc54da2e81',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
  });

  it('uses correct API url when list title option is passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
  });

  it('adds a webhook by passing list id (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'0987cfd9-f02c-479b-9fb4-3f0550462848')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: '0987cfd9-f02c-479b-9fb4-3f0550462848',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list url', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions`) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/lists/Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title including a client state', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        return {
          'clientState': 'awesome state',
          'expirationDateTime': '2019-05-29T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        clientState: 'awesome state'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'awesome state',
      'expirationDateTime': '2019-05-29T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('adds a webhook by passing list title including a expiration date', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions`) > -1) {
        return {
          'clientState': 'null',
          'expirationDateTime': '2019-01-09T23:00:00.000Z',
          'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
          'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
          'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
          'resourceData': 'null'
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2019-01-09'
      }
    });
    assert(loggerLogSpy.calledWith({
      'clientState': 'null',
      'expirationDateTime': '2019-01-09T23:00:00.000Z',
      'id': 'ef69c37d-cb0e-46d9-9758-5ebdeffd6959',
      'notificationUrl': 'https://contoso-funcions.azurewebsites.net/webhook',
      'resource': '0987cfd9-f02c-479b-9fb4-3f0550462848',
      'resourceData': 'null'
    }));
  });

  it('correctly handles a random API error', async () => {
    const errorMessage = 'An error has occurred';
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: errorMessage
          }
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2019-01-09'
      }
    } as any), new CommandError(errorMessage));
  });

  describe('schema validation (json output)', () => {
    it('fails validation when webUrl is not a SharePoint URL', () => {
      const result = schema.safeParse({
        webUrl: 'foo',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('valid SharePoint Online site URL')));
      }
    });

    it('fails validation when none of listId/listTitle/listUrl is provided', () => {
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('Specify exactly one of listId, listTitle or listUrl')));
      }
    });

    it('fails validation when multiple list identifiers are provided', () => {
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('Specify exactly one of listId, listTitle or listUrl')));
      }
    });

    it('passes validation when listTitle is provided with valid expiration', () => {
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + 10);
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook',
        expirationDateTime: futureDate.toISOString().substring(0, 10)
      });

      assert.strictEqual(result.success, true);
    });

    it('fails validation if listId is not a valid GUID', () => {
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listId: '12345',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('is not a valid GUID')));
      }
    });

    it('fails validation if expirationDateTime is not a valid date string', () => {
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook',
        expirationDateTime: '2018-X-09'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('Provide an expiration date')));
      }
    });

    it('fails validation if expirationDateTime is in the past', () => {
      const pastDate = new Date();
      pastDate.setMonth(pastDate.getMonth() - 1);
      const past = pastDate.toISOString().substring(0, 10);
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook',
        expirationDateTime: past
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('future')));
      }
    });

    it('fails validation if expirationDateTime is beyond the 6 month limit', () => {
      const futureDate = new Date();
      futureDate.setMonth(futureDate.getMonth() + 7);
      const future = futureDate.toISOString().substring(0, 10);
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook',
        expirationDateTime: future
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('future')));
      }
    });
  });
  describe('schema validation (json fallback)', () => {
    it('fails validation if expirationDateTime is not a valid date string (json output)', () => {
      const result = schema.safeParse({
        webUrl: 'https://contoso.sharepoint.com',
        listTitle: 'Documents',
        notificationUrl: 'https://contoso-functions.azurewebsites.net/webhook',
        expirationDateTime: '2018-X-09',
        output: 'json'
      });

      assert.strictEqual(result.success, false);
      if (!result.success) {
        assert(result.error.issues.some(issue => issue.message.includes('Provide an expiration date')));
      }
    });
  });

});
