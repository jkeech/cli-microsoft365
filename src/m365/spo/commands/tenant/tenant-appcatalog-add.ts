import commands from '../../commands';
import SpoCommand from '../../../base/SpoCommand';
import GlobalOptions from '../../../../GlobalOptions';
import Utils from '../../../../Utils';
import * as spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
import * as spoSiteGetCommand from '../site/site-get';
import * as spoSiteRemoveCommand from '../site/site-remove';
import * as spoSiteClassicAddCommand from '../site/site-classic-add';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as chalk from 'chalk';
import { CommandInstance } from '../../../../cli';

export interface CommandArgs {
  options: Options;
}

export interface Options extends GlobalOptions {
  force: boolean;
  owner: string;
  timeZone: string | number;
  url: string;
  wait: boolean;
}

class SpoTenantAppCatalogAddCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOG_ADD;
  }

  public get description(): string {
    return 'Creates new tenant app catalog site';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    if (this.verbose) {
      cmd.log('Checking for existing app catalog URL...');
    }

    Utils
      .executeCommandWithOutput(spoTenantAppCatalogUrlGetCommand as Command, {}, cmd)
      .then((appCatalogUrl: string): Promise<void> => {
        if (!appCatalogUrl) {
          if (this.verbose) {
            cmd.log('No app catalog URL found');
          }

          return Promise.resolve();
        }

        if (this.verbose) {
          cmd.log(`Found app catalog URL ${appCatalogUrl}`);
        }

        return this.ensureNoExistingSite(appCatalogUrl, args.options.force, cmd);
      })
      .then(_ => this.ensureNoExistingSite(args.options.url, args.options.force, cmd))
      .then(_ => this.createAppCatalog(args.options, cmd))
      .then(_ => {
        if (this.verbose) {
          cmd.log(chalk.green('DONE'));
        }
        cb();
      }, (err: CommandError): void => cb(err));
  }

  private ensureNoExistingSite(url: string, force: boolean, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (err: CommandError) => void): void => {
      if (this.verbose) {
        cmd.log(`Checking if site ${url} exists...`);
      }

      const siteGetOptions = {
        url: url,
        verbose: this.verbose,
        debug: this.debug
      };
      Utils
        .executeCommandWithOutput(spoSiteGetCommand as Command, siteGetOptions, cmd)
        .then(_ => {
          if (this.verbose) {
            cmd.log(`Found site ${url}`);
          }

          if (!force) {
            return reject(new CommandError(`Another site exists at ${url}`));
          }

          if (this.verbose) {
            cmd.log(`Deleting site ${url}...`);
          }
          const siteRemoveOptions = {
            url: url,
            skipRecycleBin: true,
            wait: true,
            confirm: true,
            verbose: this.verbose,
            debug: this.debug
          }
          Utils
            .executeCommand(spoSiteRemoveCommand as Command, siteRemoveOptions, cmd)
            .then(_ => resolve(), (err: CommandError) => reject(err));
        }, (err: CommandError) => {
          if (err.message !== 'File Not Found.' && err.message !== '404 FILE NOT FOUND') {
            // some other error occurred
            return reject(err);
          }

          if (this.verbose) {
            cmd.log(`No site found at ${url}`);
          }

          // site not found. continue
          resolve();
        });
    });
  }

  private createAppCatalog(options: Options, cmd: CommandInstance): Promise<void> {
    if (this.verbose) {
      cmd.log(`Creating app catalog at ${options.url}...`);
    }

    const siteClassicAddOptions = {
      webTemplate: 'APPCATALOG#0',
      title: 'App catalog',
      url: options.url,
      timeZone: options.timeZone,
      owner: options.owner,
      wait: options.wait,
      verbose: this.verbose,
      debug: this.debug
    };
    return Utils.executeCommand(spoSiteClassicAddCommand as Command, siteClassicAddOptions, cmd);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --url <url>',
        description: 'The absolute site url'
      },
      {
        option: '--owner <owner>',
        description: 'The account name of the site owner'
      },
      {
        option: '-z, --timeZone <timeZone>',
        description: 'Integer representing time zone to use for the site'
      },
      {
        option: '--wait',
        description: 'Wait for the site to be provisioned before completing the command'
      },
      {
        option: '--force',
        description: 'Force creating a new app catalog site if one already exists'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (typeof args.options.timeZone !== 'number') {
        return `${args.options.timeZone} is not a number`;
      }

      return true;
    };
  }
}

module.exports = new SpoTenantAppCatalogAddCommand();