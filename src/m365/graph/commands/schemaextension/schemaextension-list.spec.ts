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
import command from './schemaextension-list.js';

describe(commands.SCHEMAEXTENSION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
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
    assert.strictEqual(command.name, commands.SCHEMAEXTENSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists schema extensions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_exo2",
              "description": "sample desccription",
              "targetTypes": [
                "Message"
              ],
              "status": "Available",
              "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
              "properties": [
                {
                  "name": "p1",
                  "type": "String"
                },
                {
                  "name": "p2",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {}
    });
    try {
      assert(loggerLogSpy.calledWith([{
        "id": "adatumisv_exo2",
        "description": "sample desccription",
        "targetTypes": [
          "Message"
        ],
        "status": "Available",
        "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
        "properties": [
          {
            "name": "p1",
            "type": "String"
          },
          {
            "name": "p2",
            "type": "String"
          }
        ]
      }]
      ));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('lists two schema extensions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_exo2",
              "description": "sample desccription",
              "targetTypes": [
                "Message"
              ],
              "status": "Available",
              "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
              "properties": [
                {
                  "name": "p1",
                  "type": "String"
                },
                {
                  "name": "p2",
                  "type": "String"
                }
              ]
            },
            {
              "id": "adatumisv_exo3",
              "description": "sample desccription",
              "targetTypes": [
                "Message"
              ],
              "status": "Available",
              "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
              "properties": [
                {
                  "name": "p1",
                  "type": "String"
                },
                {
                  "name": "p2",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {}
    });
    try {
      assert(loggerLogSpy.lastCall.args[0][1].id === 'adatumisv_exo3');
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('lists schema extensions with filter options', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`$filter`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_courses",
              "description": "Extension description",
              "targetTypes": [
                "User",
                "Group"
              ],
              "status": "Available",
              "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
              "properties": [
                {
                  "name": "id",
                  "type": "Integer"
                },
                {
                  "name": "name",
                  "type": "String"
                },
                {
                  "name": "type",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        owner: '07d21ad2-c8f9-4316-a14a-347db702bd3c'
      }
    });
    try {
      assert(loggerLogSpy.calledWith([
        {
          "id": "adatumisv_courses",
          "description": "Extension description",
          "targetTypes": [
            "User",
            "Group"
          ],
          "status": "Available",
          "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
          "properties": [
            {
              "name": "id",
              "type": "Integer"
            },
            {
              "name": "name",
              "type": "String"
            },
            {
              "name": "type",
              "type": "String"
            }
          ]
        }
      ]
      ));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('lists schema extensions on the second page no page size given', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`$top`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_courses",
              "description": "Extension description",
              "targetTypes": [
                "User",
                "Group"
              ],
              "status": "Available",
              "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
              "properties": [
                {
                  "name": "id",
                  "type": "Integer"
                },
                {
                  "name": "name",
                  "type": "String"
                },
                {
                  "name": "type",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        pageNumber: 1
      }
    });
    try {
      assert(loggerLogSpy.calledWith([
        {
          "id": "adatumisv_courses",
          "description": "Extension description",
          "targetTypes": [
            "User",
            "Group"
          ],
          "status": "Available",
          "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
          "properties": [
            {
              "name": "id",
              "type": "Integer"
            },
            {
              "name": "name",
              "type": "String"
            },
            {
              "name": "type",
              "type": "String"
            }
          ]
        }
      ]
      ));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('lists schema extensions on the page size 1 second page', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`$top`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgGAAAAAAAAAA%3d%3d%23RT%3a2%23TRC%3a2%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKAAaAIcAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_courses",
              "description": "Extension description",
              "targetTypes": [
                "User",
                "Group"
              ],
              "status": "Available",
              "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
              "properties": [
                {
                  "name": "id",
                  "type": "Integer"
                },
                {
                  "name": "name",
                  "type": "String"
                },
                {
                  "name": "type",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        pageNumber: 1,
        pageSize: 1
      }
    });
    try {
      assert(loggerLogSpy.calledWith([
        {
          "id": "adatumisv_courses",
          "description": "Extension description",
          "targetTypes": [
            "User",
            "Group"
          ],
          "status": "Available",
          "owner": "07d21ad2-c8f9-4316-a14a-347db702bd3c",
          "properties": [
            {
              "name": "id",
              "type": "Integer"
            },
            {
              "name": "name",
              "type": "String"
            },
            {
              "name": "type",
              "type": "String"
            }
          ]
        }
      ]
      ));
    }
    finally {
      sinonUtil.restore(request.get);
    }
  });
  it('lists schema extensions(debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`schemaExtensions`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#schemaExtensions(*)",
          "@odata.nextLink": "https://graph.microsoft.com/v1.0/schemaExtensions?$select=*&$top=1&$skiptoken=%7B%22token%22%3a%22%2bRID%3a~F7weALI27DgBAAAAAAAAAA%3d%3d%23RT%3a1%23TRC%3a1%23ISV%3a2%23IEO%3a65551%23QCF%3a1%23FPC%3aAgEAAADKACLAQgAg2BDELgAUgQAKAgRAAIDAAAYEAEWCgACY4BEAKwSQBegLBqhBAKAACEACCAAQAAAIsAGCMQQCAgAMAgiAJaACwAQfgGqADMAFIIAAJgYAoB4AYAAAACAxBwAAAEA4EAAyACEAIABGAGAAELBAiAkIBPGAADEAABEpAAKAAAgABDKAACBMJBAgARCIIBIACQgIwBiAD8AwAAEUgQgAAAhkfAADAAAAgBCAAg0ABQCgYQAMeAIiAACgXQARAECAEIAGgAuAOYA%3d%22%2c%22range%22%3a%7B%22min%22%3a%22%22%2c%22max%22%3a%2205C1DFFFFFFFFC%22%7D%7D",
          "value": [
            {
              "id": "adatumisv_exo2",
              "description": "sample desccription",
              "targetTypes": [
                "Message"
              ],
              "status": "Available",
              "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
              "properties": [
                {
                  "name": "p1",
                  "type": "String"
                },
                {
                  "name": "p2",
                  "type": "String"
                }
              ]
            }
          ]
        };
      }

      throw 'Invalid request';
    });
    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(loggerLogSpy.calledWith([
      {
        "id": "adatumisv_exo2",
        "description": "sample desccription",
        "targetTypes": [
          "Message"
        ],
        "status": "Available",
        "owner": "617720dc-85fc-45d7-a187-cee75eaf239e",
        "properties": [
          {
            "name": "p1",
            "type": "String"
          },
          {
            "name": "p2",
            "type": "String"
          }
        ]
      }
    ]
    ));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').callsFake(async () => { throw errorMessage; });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });

  it('passes validation if the owner is a valid GUID', async () => {
    const actual = await command.validate({ options: { owner: '68be84bf-a585-4776-80b3-30aa5207aa22' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
  it('fails validation if the owner is not a valid GUID', async () => {
    const actual = await command.validate({ options: { owner: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the status is not a valid status', async () => {
    const actual = await command.validate({ options: { status: 'test' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the status is a valid status', async () => {
    const actual = await command.validate({ options: { status: 'InDevelopment' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
  it('fails validation if the pageNumber is not positive number', async () => {
    const actual = await command.validate({ options: { pageNumber: '-1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the pageNumber is a positive number', async () => {
    const actual = await command.validate({ options: { pageNumber: '2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
  it('fails validation if the pageSize is not positive number', async () => {
    const actual = await command.validate({ options: { pageSize: '-1' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
  it('passes validation if the pageSize is a positive number', async () => {
    const actual = await command.validate({ options: { pageSize: '2' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
