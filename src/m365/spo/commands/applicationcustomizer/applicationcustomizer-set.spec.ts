import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './applicationcustomizer-set.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.APPLICATIONCUSTOMIZER_SET, () => {
  let commandInfo: CommandInfo;
  const webUrl = 'https://contoso.sharepoint.com';
  const id = '14125658-a9bc-4ddf-9c75-1b5767c9a337';
  const clientSideComponentId = '015e0fcf-fe9d-4037-95af-0a4776cdfbb4';
  const title = 'SiteGuidedTour';
  const newTitle = 'New Title';
  const clientSideComponentProperties = '{"testMessage":"Updated message"}';
  let log: any[];
  let logger: Logger;

  const singleResponse = {
    value: [
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "Id": id,
        "ImageUrl": null,
        "Location": "ClientSideExtension.ApplicationCustomizer",
        "Name": title,
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": { "High": 0, "Low": 0 },
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 65536,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": "1.0.1.0"
      }
    ]
  };

  const multipleResponse = {
    value: [
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "'{testMessage:Test message}'",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": '',
        "Id": 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
        "ImageUrl": null,
        "Location": 'ClientSideExtension.ApplicationCustomizer',
        "Name": 'YourName',
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": [Object],
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": '16.0.1.0'
      },
      {
        "ClientSideComponentId": clientSideComponentId,
        "ClientSideComponentProperties": "'{testMessage:Test message}'",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": '',
        "Id": '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
        "ImageUrl": null,
        "Location": 'ClientSideExtension.ApplicationCustomizer',
        "Name": 'YourName',
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": [Object],
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": title,
        "Url": null,
        "VersionOfUserCustomAction": '16.0.1.0'
      }
    ]
  };

  const defaultUpdateCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions('${id}')`)) {
        return;
      }

      throw `Invalid request`;
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
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
      request.get,
      request.post,
      cli.handleMultipleResultsFound,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.APPLICATIONCUSTOMIZER_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, newTitle: newTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if at least one of the parameters has a value', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when all parameters are empty', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: webUrl, id: null, clientSideComponentId: null, title: '', newTitle: newTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, clientSideComponentId: 'invalid', newTitle: newTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: 'invalid', newTitle: newTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope option is not a valid scope', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, id: id, scope: 'invalid', newTitle: newTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('handles error when no application customizer with the specified id found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        return { "odata.null": true };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { id: id, webUrl: webUrl, newTitle: newTitle }
      }
      ), new CommandError(`No application customizer with id '${id}' found`));
  });

  it('handles error when no application customizer with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { title: title, webUrl: webUrl, newTitle: newTitle }
      }
      ), new CommandError(`No application customizer with title '${title}' found`));
  });

  it('handles error when no application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, newTitle: newTitle }
      }
      ), new CommandError(`No application customizer with ClientSideComponentId '${clientSideComponentId}' found`));
  });

  it('handles error when multiple application customizer with the specified title found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { title: title, webUrl: webUrl, scope: 'Site', newTitle: newTitle }
      }
      ), new CommandError("Multiple application customizer with title 'SiteGuidedTour' found. Found: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59."));
  });

  it('handles error when multiple application customizer with the specified clientSideComponentId found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    await assert.rejects(
      command.action(logger, {
        options: { clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Site', newTitle: newTitle }
      }
      ), new CommandError("Multiple application customizer with ClientSideComponentId '015e0fcf-fe9d-4037-95af-0a4776cdfbb4' found. Found: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59."));
  });

  it('handles selecting single result when multiple application customizers with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(title)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`)) {
        return multipleResponse;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(singleResponse.value[0]);

    const updateCallsSpy: sinon.SinonStub = defaultUpdateCallsStub();
    await command.action(logger, { options: { verbose: true, title: title, webUrl: webUrl, scope: 'Site', newTitle: newTitle } } as any);
    assert(updateCallsSpy.calledOnce);
  });

  it('should update the application customizer from the site by its ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        return singleResponse.value[0];
      }
      throw 'Invalid request';
    });

    const updateCallsSpy: sinon.SinonStub = defaultUpdateCallsStub();
    await command.action(logger, { options: { verbose: true, id: id, webUrl: webUrl, scope: 'Web', newTitle: newTitle } } as any);
    assert(updateCallsSpy.calledOnce);
  });

  it('should update the application customizer from the site collection by its ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${id}')`)) {
        const response = singleResponse.value[0];
        response.Scope = 2;
        return response;
      }
      throw 'Invalid request';
    });

    const updateCallsSpy: sinon.SinonStub = defaultUpdateCallsStub();
    await command.action(logger, { options: { verbose: true, id: id, webUrl: webUrl, scope: 'Site', newTitle: newTitle } } as any);
    assert(updateCallsSpy.calledOnce);
  });

  it('should update the application customizer from the site by its clientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return singleResponse;
      }
      else if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(clientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ApplicationCustomizer'))`) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    const updateCallsSpy: sinon.SinonStub = defaultUpdateCallsStub();
    await command.action(logger, { options: { verbose: true, clientSideComponentId: clientSideComponentId, webUrl: webUrl, scope: 'Web', clientSideComponentProperties: clientSideComponentProperties } } as any);
    assert(updateCallsSpy.calledOnce);
  });
});