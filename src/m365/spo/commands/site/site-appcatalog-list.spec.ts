import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./site-appcatalog-list');

describe(commands.SITE_APPCATALOG_LIST, () => {
  const appCatalogResponseValue = [
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site1",
      "ErrorMessage": null,
      "SiteID": "9798e615-b586-455e-8486-84913f492c49"
    },
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site2",
      "ErrorMessage": null,
      "SiteID": "686fe33a-7418-4a6b-92c9-d6170b1e3ae0"
    },
    {
      "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site3",
      "ErrorMessage": "Success",
      "SiteID": "2f9fd04d-2674-40ca-9ad8-d7f982dce5d0"
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      spo.getWeb
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPCATALOG_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['AbsoluteUrl', 'SiteID']);
  });

  it('retrieves site collection app catalogs within the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(appCatalogResponseValue));
  });

  it('retrieves site collection app catalogs within the tenant and exclude inaccessible sites', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getWeb').resolves();

    await command.action(logger, { options: { verbose: true, excludeDeletedSites: true } });
    assert(loggerLogSpy.calledWith(appCatalogResponseValue));
  });

  it('correctly handles error when retrieving site collection app catalogs', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        throw { error: { error: { message: 'Something went wrong' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Something went wrong'));
  });

  it('correctly handles error when retrieving site collection app catalogs and excluding inaccessible sites with 404 status', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getWeb').rejects({ status: 404, statusText: 'Not Found' });

    await command.action(logger, { options: { excludeDeletedSites: true, debug: true } });

    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles error when retrieving site collection app catalogs and excluding inaccessible sites with 403 status', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getWeb').rejects({ status: 403, statusText: 'Forbidden' });

    await command.action(logger, { options: { excludeDeletedSites: true, debug: true } });

    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles unexpected error when retrieving site collection app catalogs and excluding inaccessible sites', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites') {
        return { value: appCatalogResponseValue };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getWeb').rejects({ status: 500, statusText: 'Internal Server Error' });

    await assert.rejects(command.action(logger, { options: { excludeDeletedSites: true, debug: true } }));
  });
});