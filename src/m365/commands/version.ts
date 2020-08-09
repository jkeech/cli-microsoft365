import commands from './commands';
import AnonymousCommand from '../base/AnonymousCommand';
const packageJSON = require('../../../package.json');


class VersionCommand extends AnonymousCommand {
  public get name(): string {
    return commands.VERSION;
  }

  public get description(): string {
    return 'Shows CLI for Microsoft 365 version';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: (err?: any) => void): void {
    cmd.log(`v${packageJSON.version}`);
    cb();
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    log(vorpal.find(commands.STATUS).helpInformation());
    log(
      `  Examples:
  
    Show the version of CLI for Microsoft 365
      ${commands.STATUS}
`);
  }
}

module.exports = new VersionCommand();