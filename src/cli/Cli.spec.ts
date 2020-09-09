import * as sinon from 'sinon';
import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as chalk from 'chalk';
import * as markshell from 'markshell';
import { Cli } from '.';
import Utils from '../Utils';
import Command, { CommandOption, CommandValidate } from '../Command';
import { CommandInstance } from './CommandInstance';
const packageJSON = require('../../package.json');

class MockCommand extends Command {
  public get name(): string {
    return 'cli mock';
  }
  public get description(): string {
    return 'Mock command'
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-x, --parameterX <parameterX>',
        description: 'Required parameter'
      },
      {
        option: '-y, --parameterY [parameterY]',
        description: 'Optional parameter'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  public commandAction(cmd: CommandInstance, args: any, cb: () => void): void {
    cb();
  }
}

class MockCommandWithAlias extends Command {
  public get name(): string {
    return 'cli mock alias';
  }
  public get description(): string {
    return 'Mock command with alias'
  }
  public alias(): string[] {
    return ['cli mock alt'];
  }
  public commandAction(cmd: CommandInstance, args: any, cb: () => void): void {
    cb();
  }
}

class MockCommandWithValidation extends Command {
  public get name(): string {
    return 'cli mock validation';
  }
  public get description(): string {
    return 'Mock command with validation'
  }
  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-x, --parameterX <parameterX>',
        description: 'Required parameter'
      },
      {
        option: '-y, --parameterY [parameterY]',
        description: 'Optional parameter'
      }
    ];
    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }
  public validate(): CommandValidate {
    return (args: any): boolean | string => {
      return true;
    }
  }
  public commandAction(cmd: CommandInstance, args: any, cb: () => void): void {
    cb();
  }
}

describe('Cli', () => {
  let cli: Cli;
  let rootFolder: string;
  let cliLogStub: sinon.SinonStub;
  let cliErrorStub: sinon.SinonStub;
  let processExitStub: sinon.SinonStub;
  let markshellStub: sinon.SinonStub;
  let mockCommandActionSpy: sinon.SinonSpy;
  let mockCommand: Command;
  let mockCommandWithAlias: Command;
  let mockCommandWithValidation: Command;

  before(() => {
    cliLogStub = sinon.stub((Cli as any), 'log');
    cliErrorStub = sinon.stub((Cli as any), 'error');
    processExitStub = sinon.stub(process, 'exit');
    markshellStub = sinon.stub(markshell, 'toRawContent');

    mockCommand = new MockCommand();
    mockCommandWithAlias = new MockCommandWithAlias();
    mockCommandWithValidation = new MockCommandWithValidation();
    mockCommandActionSpy = sinon.spy(mockCommand, 'action');

    return new Promise((resolve) => {
      fs.realpath(__dirname, (err: NodeJS.ErrnoException | null, resolvedPath: string): void => {
        rootFolder = resolvedPath;
        resolve();
      });
    })
  });

  beforeEach(() => {
    cli = Cli.getInstance();
    (cli as any).loadCommand(mockCommand);
    (cli as any).loadCommand(mockCommandWithAlias);
    (cli as any).loadCommand(mockCommandWithValidation);
  });

  afterEach(() => {
    (cli as any).instance = undefined;
    cliLogStub.reset();
    cliErrorStub.reset();
    processExitStub.reset();
    markshellStub.reset();
    mockCommandActionSpy.resetHistory();
    Utils.restore([
      Cli.executeCommand,
      fs.existsSync,
      mockCommandWithValidation.validate,
      mockCommandWithValidation.action
    ]);
  });

  after(() => {
    Utils.restore([
      (Cli as any).log,
      (Cli as any).error,
      process.exit,
      markshell.toRawContent
    ]);
  });

  it('shows generic help when no command specified', (done) => {
    cli
      .execute(rootFolder, [])
      .then(_ => {
        try {
          assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('exits with 0 code when no command specified', (done) => {
    cli
      .execute(rootFolder, [])
      .then(_ => {
        try {
          assert(processExitStub.calledWith(0));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows generic help when help command and no command name specified', (done) => {
    cli
      .execute(rootFolder, ['help'])
      .then(_ => {
        try {
          assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows generic help when --help option specified', (done) => {
    cli
      .execute(rootFolder, ['--help'])
      .then(_ => {
        try {
          assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows generic help when -h option specified', (done) => {
    cli
      .execute(rootFolder, ['-h'])
      .then(_ => {
        try {
          assert(cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows help for the specific command when help specified followed by a valid command name', (done) => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.md'));
    cli
      .execute(rootFolder, ['help', 'cli', 'mock'])
      .then(_ => {
        try {
          assert(markshellStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows help for the specific command when valid command name specified followed by --help', (done) => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.md'));
    cli
      .execute(rootFolder, ['cli', 'mock', '--help'])
      .then(_ => {
        try {
          assert(markshellStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows help for the specific command when valid command name specified followed by -h', (done) => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.md'));
    cli
      .execute(rootFolder, ['cli', 'mock', '-h'])
      .then(_ => {
        try {
          assert(markshellStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it('shows help for the specific command when help specified followed by a valid command alias', (done) => {
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().endsWith('.md'));
    cli
      .execute(rootFolder, ['help', 'cli', 'mock', 'alt'])
      .then(_ => {
        try {
          assert(cliLogStub.called);
          assert(!cliLogStub.calledWith(`CLI for Microsoft 365 v${packageJSON.version}`));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`passes options validation if the command doesn't allow unknown options and specified options match command options`, (done) => {
    cli
      .execute(rootFolder, ['cli', 'mock', '-x', '123', '-y', '456'])
      .then(_ => {
        try {
          assert(mockCommandActionSpy.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`fails options validation if the command doesn't allow unknown options and specified options match command options`, (done) => {
    cli
      .execute(rootFolder, ['cli', 'mock', '-x', '123', '-z'])
      .then(_ => {
        try {
          assert(cliErrorStub.calledWith(chalk.red(`Error: Invalid option: 'z'${os.EOL}`)));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`doesn't execute command action when option validation failed`, (done) => {
    cli
      .execute(rootFolder, ['cli', 'mock', '-x', '123', '-z'])
      .then(_ => {
        try {
          assert(mockCommandActionSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`exits with exit code 1 when option validation failed`, (done) => {
    cli
      .execute(rootFolder, ['cli', 'mock', '-x', '123', '-z'])
      .then(_ => {
        try {
          assert(processExitStub.calledWith(1));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`fails validation if a required option is missing`, (done) => {
    cli
      .execute(rootFolder, ['cli', 'mock'])
      .then(_ => {
        try {
          assert(cliErrorStub.calledWith(chalk.red(`Error: Required option parameterX not specified`)));
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`calls command's validation method when defined`, (done) => {
    const mockCommandValidateSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'validate');
    cli
      .execute(rootFolder, ['cli', 'mock', 'validation', '-x', '123'])
      .then(_ => {
        try {
          assert(mockCommandValidateSpy.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`passes validation when the command's validate method returns true`, (done) => {
    sinon.stub(mockCommandWithValidation, 'validate').callsFake(() => () => true);
    const mockCommandWithValidationActionSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'action');

    cli
      .execute(rootFolder, ['cli', 'mock', 'validation', '-x', '123'])
      .then(_ => {
        try {
          assert(mockCommandWithValidationActionSpy.called);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });

  it(`fails validation when the command's validate method returns a string`, (done) => {
    sinon.stub(mockCommandWithValidation, 'validate').callsFake(() => () => 'Error');
    const mockCommandWithValidationActionSpy: sinon.SinonSpy = sinon.spy(mockCommandWithValidation, 'action');

    cli
      .execute(rootFolder, ['cli', 'mock', 'validation', '-x', '123'])
      .then(_ => {
        try {
          assert(mockCommandWithValidationActionSpy.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => done(e));
  });
});