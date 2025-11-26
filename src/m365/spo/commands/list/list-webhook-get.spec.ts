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
import command from './list-webhook-get.js';

describe(commands.LIST_WEBHOOK_GET, () => {
  const webhookGetResponse = {
    "clientState": null,
    "expirationDateTime": "2019-01-27T16:32:05.4610008Z",
    "id": "cc27a922-8224-4296-90a5-ebbc54da2e85",
    "notificationUrl": "https://mlk-document-publishing-fa-dev-we.azurewebsites.net/api/HandleWebHookNotification?code=jZyDfmBffPn7x0xYCQtZuxfqapu7cJzJo6puvruJiMUOxUl6XkxXAA==",
    "resource": "dfddade1-4729-428d-881e-7fedf3cae50d",
    "resourceData": null
  };
  const webUrl = 'https://contoso.sharepoint.com';
  const listId = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';
  const webhookId = 'cc27a922-8224-4296-90a5-ebbc54da2e85';
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_WEBHOOK_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves specified webhook of the given list if title option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists/GetByTitle('Documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listTitle: 'Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    });
    assert(loggerLogSpy.calledWith(webhookGetResponse));
  });

  it('retrieves specified webhook of the given list if url option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('${formatting.encodeQueryParameter('/sites/ninja/lists/Documents')}')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/lists/Documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    });
    assert(loggerLogSpy.calledWith(webhookGetResponse));
  });

  it('retrieves specific webhook of the specific list if id option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(webhookGetResponse));
  });

  it('retrieves specific webhook of the specific list if url option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/ninja/_api/web/GetList('%2Fsites%2Fninja%2Fshared%20documents')/Subscriptions('cc27a922-8224-4296-90a5-ebbc54da2e85')`) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return webhookGetResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listUrl: '/sites/ninja/shared documents',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85',
        verbose: true
      }
    });
    assert(loggerLogSpy.calledWith(webhookGetResponse));
  });


  it('correctly handles error when getting information for a site that doesn\'t exist', async () => {
    const error = {
      error: {
        'odata.error': {
          code: '-1, Microsoft.SharePoint.Client.InvalidOperationException',
          message: {
            value: '404 - File not found'
          }
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`https://contoso.sharepoint.com/sites/ninja/_api/web/lists(guid'dfddade1-4729-428d-881e-7fedf3cae50d')/Subscriptions('ab27a922-8224-4296-90a5-ebbc54da1981')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          throw error;
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'ab27a922-8224-4296-90a5-ebbc54da1981'
      }
    } as any), new CommandError(error.error['odata.error'].message.value));
  });

  it('command correctly handles list get reject request', async () => {
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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        throw error;
      }

      throw 'Invalid request';
    });

    const actionTitle: string = 'Documents';

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        listTitle: actionTitle,
        webUrl: 'https://contoso.sharepoint.com/sites/ninja'
      }
    }), new CommandError(error.error['odata.error'].message.value));
  });

  it('uses correct API url when id option is passed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return 'Correct Url';
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/ninja',
        listId: 'dfddade1-4729-428d-881e-7fedf3cae50d',
        id: 'cc27a922-8224-4296-90a5-ebbc54da2e85'
      }
    });
  });

  it('passes validation when listId is provided', () => {
    const result = schema.safeParse({ webUrl, listId, id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when listTitle is provided', () => {
    const result = schema.safeParse({ webUrl, listTitle: 'Documents', id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when listUrl is provided', () => {
    const result = schema.safeParse({ webUrl, listUrl: '/sites/site/lists/lib', id: webhookId });
    assert.strictEqual(result.success, true);
  });

  it('fails validation when list identifier is missing', () => {
    const result = schema.safeParse({ webUrl, id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('Specify exactly one of listId, listTitle or listUrl.')));
  });

  it('fails validation when listId is not a GUID', () => {
    const result = schema.safeParse({ webUrl, listId: 'abc', id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('not a valid GUID')));
  });

  it('fails validation when id is not a GUID', () => {
    const result = schema.safeParse({ webUrl, listId, id: 'abc' });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('not a valid GUID')));
  });

  it('fails validation when webUrl is not a SharePoint URL', () => {
    const result = schema.safeParse({ webUrl: 'foo', listId, id: webhookId });
    assert(result.success === false && result.error.issues.some(issue => issue.message.includes('SharePoint Online site URL')));
  });
});
