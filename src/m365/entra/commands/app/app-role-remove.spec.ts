import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-role-remove.js';
import { settingsNames } from '../../../../settingsNames.js';
import { entraApp } from '../../../../utils/entraApp.js';

describe(commands.APP_ROLE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  //#region Mocked Responses 
  const appResponse = {
    value: [
      {
        "id": "5b31c38c-2584-42f0-aa47-657fb3a84230"
      }
    ]
  };
  //#endregion

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      entraApp.getAppRegistrationByAppId
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_ROLE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --force option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role name and --force option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'ProductRead',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role id and --force option specified', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role claim and --force option specified', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        claim: 'Product.Read',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --force option specified', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --force option specified', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and --force option specified (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and --force option specified (debug)', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and --force option specified (debug)', async () => {
    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role id and --force option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role claim and --force option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        claim: 'Product.Read',
        force: true
      }
    });
  });

  it('deletes an app role when the role is in "disabled" state and valid appId, role name and --force option specified (debug)', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    });
  });

  it('handles error when multiple apps with the specified appName found and --force option is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            },
            {
              id: 'a39c738c-939e-433b-930d-b02f2931a08b'
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        force: true
      }
    }), new CommandError(`Multiple Microsoft Entra application registrations with name 'App-Name' found. Found: 5b31c38c-2584-42f0-aa47-657fb3a84230, a39c738c-939e-433b-930d-b02f2931a08b.`));
  });

  it('handles selecting single result when multiple apps with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            },
            {
              id: 'a39c738c-939e-433b-930d-b02f2931a08b'
            }
          ]
        };
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": false,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: '5b31c38c-2584-42f0-aa47-657fb3a84230' });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          removeRequestIssued = true;
          return;
        }
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    });
    assert(removeRequestIssued);
  });

  it('handles when multiple roles with the same name are found and --force option specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product get",
              "displayName": "ProductRead",
              "id": "9267ab18-8d09-408d-8c94-834662ed16d1",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Get"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    }), new CommandError(`Multiple roles with name 'ProductRead' found. Found: c4352a0a-494f-46f9-b843-479855c173a7, 9267ab18-8d09-408d-8c94-834662ed16d1.`));
  });

  it('handles selecting single result when multiple roles with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;
    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product get",
              "displayName": "ProductRead",
              "id": "9267ab18-8d09-408d-8c94-834662ed16d1",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Get"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: 'c4352a0a-494f-46f9-b843-479855c173a7' });

    sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Get" &&
          appRole.id === '9267ab18-8d09-408d-8c94-834662ed16d1' &&
          appRole.isEnabled === true) {
          removeRequestIssued = true;
          return;
        }
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    });
    assert(removeRequestIssued);
  });

  it('handles when no roles with the specified name are found and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        name: 'ProductRead',
        force: true
      }
    }), new CommandError(`No app role with name 'ProductRead' found.`));
  });

  it('handles when no roles with the specified claim are found and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        claim: 'Product.Read',
        force: true
      }
    }), new CommandError(`No app role with claim 'Product.Read' found.`));
  });

  it('handles when no roles with the specified id are found and --force option specified', async () => {

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.onFirstCall().callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'App-Name'&$select=id`) {
        return {
          "value": [
            {
              id: '5b31c38c-2584-42f0-aa47-657fb3a84230'
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    getRequestStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: []
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'App-Name',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: true
      }
    }), new CommandError(`No app role with id 'c4352a0a-494f-46f9-b843-479855c173a7' found.`));
  });

  it('prompts before removing the specified app role when force option not passed', async () => {
    await command.action(logger, { options: { appName: 'App-Name', claim: 'Product.Read' } });

    assert(promptIssued);
  });

  it('prompts before removing the specified app role when force option not passed (debug)', async () => {
    await command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } });

    assert(promptIssued);
  });

  it('deletes an app role when the role is in enabled state and valid appObjectId, role claim and the prompt is confirmed (debug)', async () => {

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const getRequestStub = sinon.stub(request, 'get');

    getRequestStub.callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        debug: true,
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        claim: 'Product.Read',
        force: false
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role name and prompt is confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await command.action(logger, {
      options: {
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        name: 'ProductRead',
        force: false
      }
    });
  });

  it('deletes an app role when the role is in enabled state and valid appId, role id and prompt is confirmed (debug)', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    sinon.stub(entraApp, 'getAppRegistrationByAppId').resolves(appResponse.value[0]);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        return {
          id: '5b31c38c-2584-42f0-aa47-657fb3a84230',
          appRoles: [
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product read",
              "displayName": "ProductRead",
              "id": "c4352a0a-494f-46f9-b843-479855c173a7",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Read"
            },
            {
              "allowedMemberTypes": [
                "User"
              ],
              "description": "Product write",
              "displayName": "ProductWrite",
              "id": "54e8e043-86db-49bb-bfa8-c9c27ebdf3b6",
              "isEnabled": true,
              "lang": null,
              "origin": "Application",
              "value": "Product.Write"
            }
          ]
        };
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    const patchStub = sinon.stub(request, 'patch');

    patchStub.onFirstCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 2) {
        const appRole = opts.data.appRoles[0];
        if (appRole.isEnabled === false) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    patchStub.onSecondCall().callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230' &&
        opts.data &&
        opts.data.appRoles.length === 1) {
        const appRole = opts.data.appRoles[0];
        if (appRole.value === "Product.Write" &&
          appRole.id === '54e8e043-86db-49bb-bfa8-c9c27ebdf3b6' &&
          appRole.isEnabled === true) {
          return;
        }
      }
      throw `Invalid request ${JSON.stringify(opts)}`;
    });


    await command.action(logger, {
      options: {
        debug: true,
        appId: '53788d97-dc06-460c-8bd6-5cfbc7e3b0f7',
        id: 'c4352a0a-494f-46f9-b843-479855c173a7',
        force: false
      }
    });
  });

  it('aborts deleting app role when prompt is not confirmed', async () => {
    // represents the Microsoft Entra app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { appName: 'App-Name', claim: 'Product.Read' } });
    assert(patchStub.notCalled);
  });

  it('aborts deleting app role when prompt is not confirmed (debug)', async () => {
    // represents the Microsoft Entra app get request called when the prompt is confirmed
    const patchStub = sinon.stub(request, 'get');
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { debug: true, appName: 'App-Name', claim: 'Product.Read' } });
    assert(patchStub.notCalled);
  });

  it('handles error when the app specified with appObjectId not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/myorganization/applications/5b31c38c-2584-42f0-aa47-657fb3a84230?$select=id,appRoles') {
        throw {
          "error": {
            "code": "Request_ResourceNotFound",
            "message": "Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.",
            "innerError": {
              "date": "2021-04-20T17:22:30",
              "request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d",
              "client-request-id": "f58cc4de-b427-41de-b37c-46ee4925a26d"
            }
          }
        };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appObjectId: '5b31c38c-2584-42f0-aa47-657fb3a84230',
        name: 'App-Role',
        force: true
      }
    }), new CommandError(`Resource '5b31c38c-2584-42f0-aa47-657fb3a84230' does not exist or one of its queried reference-property objects are not present.`));
  });

  it('handles error when the app specified with the appId not found', async () => {
    const error = `App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`;
    sinon.stub(entraApp, 'getAppRegistrationByAppId').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f',
        name: 'App-Role',
        force: true
      }
    }), new CommandError(`App with appId '9b1b1e42-794b-4c71-93ac-5ed92488b67f' not found in Microsoft Entra ID`));
  });

  it('handles error when the app specified with appName not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/myorganization/applications?$filter=displayName eq 'My%20app'&$select=id`) {
        return { value: [] };
      }

      throw `Invalid request ${JSON.stringify(opts)}`;
    });

    await assert.rejects(command.action(logger, {
      options: {
        appName: 'My app',
        name: 'App-Role',
        force: true
      }
    }), new CommandError(`No Microsoft Entra application registration with name My app found`));
  });

  it('fails validation if appId and appObjectId specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appObjectId: 'c75be2e1-0204-4f95-857d-51a37cf40be8' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if appObjectId and appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', appName: 'My app' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither appId, appObjectId nor appName specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role name and id is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation role name and claim is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: "Product read", claim: "Product.Read" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if role id and claim is specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: "Product.Read", id: "c4352a0a-494f-46f9-b843-479855c173a7" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither role name, id or claim specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if specified role id is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '77355bee' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,name', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,claim', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appId,id', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,name', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,claim', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appObjectId,id', async () => {
    const actual = await command.validate({ options: { appObjectId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,name', async () => {
    const actual = await command.validate({ options: { appName: 'My App', name: 'ProductRead' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,claim', async () => {
    const actual = await command.validate({ options: { appName: 'My App', claim: 'Product.Read' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified - appName,id', async () => {
    const actual = await command.validate({ options: { appName: 'My App', id: '4e241a08-3a95-4c47-8c68-8c0df7d62ce2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
