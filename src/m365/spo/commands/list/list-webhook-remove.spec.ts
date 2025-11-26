import assert from 'assert';
import sinon from 'sinon';
import { z } from 'zod';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-webhook-remove.js';

describe(commands.LIST_WEBHOOK_REMOVE, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '0cd891ef-afce-4e55-b836-fce03286cccf';
  const listTitle = 'Documents';
  const listUrl = '/sites/ninja/lists/Documents';
  const webhookId = 'cc27a922-8224-4296-90a5-ebbc54da2e81';
  let log: any[];
  let logger: Logger;
  let requests: any[];
  let promptIssued = false;
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing webhook from list when confirmation argument not passed (list title)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle, id: webhookId } });

    assert(promptIssued);
  });

  it('prompts before removing webhook from list when confirmation argument not passed (list url)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl: '/sites/ninja/Documents', id: webhookId } });

    assert(promptIssued);
  });

  it('prompts before removing list when confirmation argument not passed (list id)', async () => {
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId, id: webhookId } });

    assert(promptIssued);
  });

  it('aborts removing list when prompt not confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle, id: webhookId } });
    assert(requests.length === 0);
  });

  it('removes the list (retrieved by Title) when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle, id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by Title) webhook when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listTitle, id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) > -1 &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by id) webhook when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by id) webhook when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by id) webhook when prompt confirmed in options', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listId: 'dfddade1-4729-428d-881e-7fedf3cae50d', id: webhookId, force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by url) webhook when confirmed in options', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl, id: webhookId, force: true } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by url) webhook when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl, id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('%2Fsites%2Fninja%2Flists%2FDocuments')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('removes the list (retrieved by url) webhook when prompt confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      requests.push(opts);

      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return;
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/ninja', listUrl, id: webhookId } });
    let correctRequestIssued = false;
    requests.forEach(r => {
      if (r.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e81')` &&
        r.headers.accept &&
        r.headers.accept.indexOf('application/json') === 0) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('handles error correctly', async () => {
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
    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        id: '0cd891ef-afce-4e55-b836-fce03286cccf',
        webUrl,
        listTitle,
        force: true
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('passes validation when listId is provided', () => {
    const result = schema.safeParse({ webUrl, listId, id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when listTitle is provided', () => {
    const result = schema.safeParse({ webUrl, listTitle, id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when listUrl is provided', () => {
    const result = schema.safeParse({ webUrl, listUrl, id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('fails validation when no list identifier is supplied', () => {
    const result = schema.safeParse({ webUrl, id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('Specify exactly one of listId, listTitle or listUrl.')));
  });

  it('fails validation when id is not a GUID', () => {
    const result = schema.safeParse({ webUrl, listId, id: 'abc' });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('not a valid GUID')));
  });

  it('fails validation when listId is not a GUID', () => {
    const result = schema.safeParse({ webUrl, listId: 'abc', id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('not a valid GUID')));
  });

  it('fails validation when webUrl is not a SharePoint Online site URL', () => {
    const result = schema.safeParse({ webUrl: 'foo', listId, id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('SharePoint Online site URL')));
  });
});
