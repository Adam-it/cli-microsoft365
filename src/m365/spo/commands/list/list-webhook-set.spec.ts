import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-webhook-set.js';

describe(commands.LIST_WEBHOOK_SET, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const siteUrl = 'https://contoso.sharepoint.com/sites/ninja';
  const listId = 'cc27a922-8224-4296-90a5-ebbc54da2e77';
  const listTitle = 'Documents';
  const listUrl = '/sites/ninja/lists/Documents';
  const webhookId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  let log: any[];
  let logger: Logger;
  let schema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    schema = command.getSchemaToParse()!;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url when list id option is passed', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: webhookId,
        webUrl,
        listId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
  });

  it('uses correct API url when list title option is passed', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: webhookId,
        webUrl,
        listTitle,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
  });

  it('updates notification url and expiration date of the webhook by passing list title (debug)', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09T00:00:00.000Z'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        debug: true,
        webUrl: siteUrl,
        listTitle,
        id: webhookId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates notification url and expiration date of the webhook by passing list id (verbose)', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09T00:00:00.000Z'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'cc27a922-8224-4296-90a5-ebbc54da2e77')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: siteUrl,
        listId,
        id: webhookId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates notification url and expiration date of the webhook by passing list title', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
      expirationDateTime: '2018-10-09T00:00:00.000Z'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listTitle,
        id: webhookId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook',
        expirationDateTime: '2018-10-09'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates notification url of the webhook by passing list title', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listTitle,
        id: webhookId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates notification url of the webhook by passing list url', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listUrl,
        id: webhookId,
        notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates clientState of the webhook by passing list url', async () => {
    let actual: string = '';
    const clientState = 'My client state';
    const expected: string = JSON.stringify({
      clientState: clientState
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listUrl,
        id: webhookId,
        clientState: clientState
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates expiration date of the webhook by passing list title', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      expirationDateTime: '2019-03-02T00:00:00.000Z'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listTitle,
        id: webhookId,
        expirationDateTime: '2019-03-02'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('updates expiration date of the webhook by passing list url', async () => {
    let actual: string = '';
    const expected: string = JSON.stringify({
      expirationDateTime: '2019-03-02T00:00:00.000Z'
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        actual = JSON.stringify(opts.data);
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        verbose: true,
        webUrl: siteUrl,
        listUrl,
        id: webhookId,
        expirationDateTime: '2019-03-02'
      }
    });
    assert.strictEqual(actual, expected);
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: 'An error has occurred'
          }
        }
      }
    };
    sinon.stub(request, 'patch').rejects(error);

    await assert.rejects(command.action(logger, {
      options:
      {
        webUrl: siteUrl,
        listTitle,
        id: webhookId,
        expirationDateTime: '2019-03-02'
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const result = schema.safeParse({
      webUrl: 'foo',
      listTitle,
      id: webhookId,
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('SharePoint Online site URL')));
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const result = schema.safeParse({
      webUrl,
      listId,
      id: webhookId,
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(result.success, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const result = schema.safeParse({
      webUrl,
      listId,
      id: '12345',
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('valid GUID')));
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const result = schema.safeParse({
      webUrl,
      listId: '12345',
      id: webhookId,
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('valid GUID')));
  });

  it('fails validation if notificationUrl, expirationDateTime or clientState options are not passed', () => {
    const result = schema.safeParse({
      webUrl,
      listTitle,
      id: webhookId
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('at least one is required')));
  });

  it('fails validation if multiple list identifiers are provided', () => {
    const result = schema.safeParse({
      webUrl,
      listId,
      listTitle,
      id: webhookId,
      notificationUrl: 'https://contoso-funcions.azurewebsites.net/webhook'
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('exactly one')));
  });

  it('fails validation if the expirationDateTime option is not a valid date string', () => {
    const result = schema.safeParse({
      webUrl,
      listTitle,
      id: webhookId,
      expirationDateTime: '2018-X-09'
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('expiration date which is a date time in the future')));
  });

  it('fails validation if the expirationDateTime option is beyond the 6 month window', () => {
    const future = new Date();
    future.setFullYear(future.getFullYear() + 1);
    const result = schema.safeParse({
      webUrl,
      listTitle,
      id: webhookId,
      expirationDateTime: future.toISOString()
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('within 6 months')));
  });
});
