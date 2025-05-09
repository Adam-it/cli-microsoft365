import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './term-group-add.js';

describe(commands.TERM_GROUP_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TERM_GROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds term group by name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.data.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { name: 'PnPTermSets' } });
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
      "Description": ""
    }));
  });

  it('adds term group by id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8' } } as any);
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a8",
      "Description": ""
    }));
  });

  it('adds term group by name with description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.data.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { name: 'PnPTermSets', description: 'Term sets for PnP' } });
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
      "Description": "Term sets for PnP"
    }));
  });

  it('adds term group by id with description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8', description: 'Term sets for PnP' } } as any);
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a8",
      "Description": "Term sets for PnP"
    }));
  });

  it('adds term group by id with description when webUrl is specified', async () => {
    const webUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8', description: 'Term sets for PnP', webUrl: webUrl } } as any);
    assert(loggerLogSpy.calledWith({ "Name": "PnPTermSets", "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a8", "Description": "Term sets for PnP" }));
  });

  it('correctly handles error when retrieving the term store', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: 'PnPTermSets' } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles error when the specified name already exists', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.data.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          throw "Group names must be unique.";
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: 'PnPTermSets' } } as any), new CommandError('Group names must be unique.'));
  });

  it('correctly handles error when the specified id already exists', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'PnPTermSets',
        id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8'
      }
    } as any), new CommandError('Failed to read from or write to database. Refresh and try again. If the problem persists, please contact the administrator.'));
  });

  it('correctly handles error when setting the description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{6cb612c7-2e96-47b9-b7c7-41ddc87379a8}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a8)\/", "Description": ""
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": {
                "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "304b919e-c041-0000-29c7-027259fd7cb6"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, name: 'PnPTermSets', id: '6cb612c7-2e96-47b9-b7c7-41ddc87379a8', description: 'Term sets for PnP' } } as any), new CommandError('An error has occurred'));
  });

  it('correctly escapes XML in term group name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets&gt;</Parameter><Parameter Type="Guid">{`) > -1 && opts.data.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets>", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { name: 'PnPTermSets>' } });
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets>",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
      "Description": ""
    }));
  });

  it('correctly escapes XML in term group description', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><Query Id="9" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><StaticMethod Id="3" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="6" ParentId="3" Name="GetDefaultSiteCollectionTermStore" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-5076-0000-29c7-0729bb255d69"
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:ss:"
            }, 7, {
              "IsNull": false
            }, 8, {
              "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ=="
            }, 9, {
              "_ObjectType_": "SP.Taxonomy.TermStore", "_ObjectIdentity_": "d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==", "DefaultLanguage": 1033, "Id": "\/Guid(707e4d61-bd1c-4bc1-a1fd-fce01591a951)\/", "IsOnline": true, "Languages": [
                1033
              ], "Name": "Taxonomy_tDB2pT87w98nSRfEZAp8tQ==", "WorkingLanguage": 1033
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="Name" ScalarProperty="true" /><Property Name="Id" ScalarProperty="true" /><Property Name="Description" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="13" ParentId="6" Name="CreateGroup"><Parameters><Parameter Type="String">PnPTermSets</Parameter><Parameter Type="Guid">{`) > -1 && opts.data.indexOf(`}</Parameter></Parameters></Method><Identity Id="6" Name="d94a919e-5076-0000-29c7-0729bb255d69|fec14c62-7c3b-481b-851b-c80d7802b224:st:YU1+cBy9wUuh\u002ffzgFZGpUQ==" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "d94a919e-0083-0000-29c7-00c65c41f487"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac="
            }, 16, {
              "_ObjectType_": "SP.Taxonomy.TermGroup", "_ObjectIdentity_": "d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=", "Name": "PnPTermSets", "Id": "\/Guid(6cb612c7-2e96-47b9-b7c7-41ddc87379a7)\/", "Description": ""
            }
          ]);
        }

        if (opts.data.indexOf(`<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="51" ObjectPathId="45" Name="Description"><Parameter Type="String">Term sets for PnP&gt;</Parameter></SetProperty></Actions><ObjectPaths><Identity Id="45" Name="d94a919e-0083-0000-29c7-00c65c41f487|fec14c62-7c3b-481b-851b-c80d7802b224:gr:YU1+cBy9wUuh\u002ffzgFZGpUccStmyWLrlHt8dB3chzeac=" /></ObjectPaths></Request>`) > -1) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8112.1217", "ErrorInfo": null, "TraceCorrelationId": "164b919e-40fa-0000-2cdb-e0b737b04e48"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { name: 'PnPTermSets', description: 'Term sets for PnP>' } });
    assert(loggerLogSpy.calledWith({
      "Name": "PnPTermSets",
      "Id": "6cb612c7-2e96-47b9-b7c7-41ddc87379a7",
      "Description": "Term sets for PnP>"
    }));
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { name: 'PnPTermSets', id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid webUrl', async () => {
    const actual = await command.validate({ options: { name: 'PnPTermSets', webUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id and name specified', async () => {
    const actual = await command.validate({ options: { name: 'PnPTermSets', id: '9e54299e-208a-4000-8546-cc4139091b26' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and webUrl are specified', async () => {
    const actual = await command.validate({ options: { name: 'People', webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
