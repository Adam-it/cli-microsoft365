import * as assert from 'assert';
import * as fs from 'fs';
import Command, { CommandError } from '../../../../Command';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { spfx } from '../../../../utils/spfx';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import sinon = require('sinon');
import path = require('path');
const command: Command = require('./project-github-workflow-add');

describe(commands.PROJECT_GITHUB_WORKFLOW_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const projectPath: string = path.resolve('/test-project');

  before(() => {
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spfx, 'getHighestNodeVersion').returns('22.0.x');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_GITHUB_WORKFLOW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if loginMethod is not valid type', async () => {
    const actual = await command.validate({ options: { loginMethod: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is not valid type', async () => {
    const actual = await command.validate({ options: { scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is sitecollection but the siteUrl was not defined', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required properties are provided', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('creates a default workflow (debug)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: { debug: true } } as any);
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.github', 'workflows', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('creates a default workflow with specifying options', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'mkdirSync').callsFake((fakePath, options) => {
      if (fakePath.toString() === path.join(projectPath, '.github') && (options as fs.MakeDirectoryOptions).recursive) {
        return path.join(projectPath, '.github');
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: { name: 'test', branchName: 'dev', manuallyTrigger: true, skipFeatureDeployment: true, overwrite: true, loginMethod: 'user', scope: 'sitecollection' } } as any);
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.github', 'workflows', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('handles error with unknown version of SPFx', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns(undefined);

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.'));

  });

  it('handles error with not found node version', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('99.99.99');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} }), new CommandError("Could not find Node version for version '99.99.99' of SharePoint Framework."));
  });

  it('handles unexpected error', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').throws(
      new Error('writeFileSync failed')
    );

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('writeFileSync failed'));
  });
});
