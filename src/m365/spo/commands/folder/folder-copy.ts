import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import * as url from 'url';
import { CommandInstance } from '../../../../cli';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  allowSchemaMismatch?: boolean;
}

interface JobProgressOptions {
  webUrl: string;
  /**
   * Response object retrieved from /_api/site/CreateCopyJobs
   */
  copyJopInfo: any;
  /**
   * Poll interval to call /_api/site/GetCopyJobProgress
   */
  progressPollInterval: number;
  /**
   * Max poll intervals to call /_api/site/GetCopyJobProgress
   * after which to give up
   */
  progressMaxPollAttempts: number;
  /**
   * Retry attempts before give up.
   * Give up if /_api/site/GetCopyJobProgress returns 
   * X reject promises in a row
   */
  progressRetryAttempts: number;
}

class SpoFolderCopyCommand extends SpoCommand {
  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const webUrl: string = args.options.webUrl;
    const parsedUrl: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${parsedUrl.protocol}//${parsedUrl.hostname}`;

    const sourceAbsoluteUrl: string = this.urlCombine(webUrl, args.options.sourceUrl);
    const allowSchemaMismatch: boolean = args.options.allowSchemaMismatch || false;
    const requestUrl: string = this.urlCombine(webUrl, '/_api/site/CreateCopyJobs');
    const requestOptions: any = {
      url: requestUrl,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      body: {
        exportObjectUris: [sourceAbsoluteUrl],
        destinationUri: this.urlCombine(tenantUrl, args.options.targetUrl),
        options: {
          "AllowSchemaMismatch": allowSchemaMismatch,
          "IgnoreVersionHistory": true
        }
      },
      json: true
    };

    request
      .post(requestOptions)
      .then((jobInfo: any): Promise<any> => {
        const jobProgressOptions: JobProgressOptions = {
          webUrl: webUrl,
          copyJopInfo: jobInfo.value[0],
          progressMaxPollAttempts: 1000, // 1 sec.
          progressPollInterval: 30 * 60, // approx. 30 mins. if interval is 1000
          progressRetryAttempts: 5
        }

        return this.waitForJobResult(jobProgressOptions, cmd);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  /**
   * A polling function that awaits the 
   * Azure queued copy job to return JobStatus = 0 meaning it is done with the task.
   */
  private waitForJobResult(opts: JobProgressOptions, cmd: CommandInstance):
    Promise<void> {

    let pollCount: number = 0;
    let retryAttemptsCount: number = 0;

    const checkCondition = (resolve: () => void, reject: (error: any) => void): void => {
      pollCount++;
      const requestUrl: string = `${opts.webUrl}/_api/site/GetCopyJobProgress`;
      const requestOptions: any = {
        url: requestUrl,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        body: { "copyJobInfo": opts.copyJopInfo },
        json: true
      };

      request.post(requestOptions).then((resp: any): void => {
        retryAttemptsCount = 0; // clear retry on promise success 

        if (this.verbose) {
          if (resp.JobState && resp.JobState === 4) {
            cmd.log(`Check #${pollCount}. Copy job in progress... JobState: ${resp.JobState}`);
          }
          else {
            cmd.log(`Check #${pollCount}. JobState: ${resp.JobState}`);
          }
        }

        for (const item of resp.Logs) {
          const log = JSON.parse(item);

          // reject if progress error 
          if (log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // three possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        // max poll attempts flag raised = reject promise returned
        if (resp.JobState === 0) {
          // job done
          resolve();
        }
        else if (pollCount < opts.progressMaxPollAttempts) {
          // if the condition isn't met but the timeout hasn't elapsed, go again
          setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
        }
        else {
          reject(new Error('waitForJobResult timed out'));
        }
      },
        (error: any) => {
          retryAttemptsCount++;

          // let's retry x times in row before we give up since
          // this is progress check and even if rejects a promise
          // the actual copy process can success.
          if (retryAttemptsCount <= opts.progressRetryAttempts) {
            setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
          }
          else {
            reject(error);
          }
        });
    };

    return new Promise<void>(checkCondition);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder is located'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the folder to copy'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to copy the folder'
      },
      {
        option: '--allowSchemaMismatch',
        description: 'Ignores any missing fields in the target document library and copies the folder anyway'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }
}

module.exports = new SpoFolderCopyCommand();