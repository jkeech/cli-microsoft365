import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';


class SpoReportActivityFileCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.REPORT_ACTIVITYFILECOUNTS;
  }

  public get usageEndpoint(): string {
    return 'getSharePointActivityFileCounts';
  }

  public get description(): string {
    return 'Gets the number of unique, licensed users who interacted with files stored on SharePoint sites';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the number of unique, licensed users who interacted with files stored
    on SharePoint sites for the last week
      m365 ${this.name} --period D7

    Gets the number of unique, licensed users who interacted with files stored
    on SharePoint sites for the last week and exports the report data
    in the specified path in text format
      m365 ${this.name} --period D7 --output text > "activityfilecounts.txt"

    Gets the number of unique, licensed users who interacted with files stored
    on SharePoint sites for the last week and exports the report data
    in the specified path in json format
      m365 ${this.name} --period D7 --output json > "activityfilecounts.json"
`);
  }
}

module.exports = new SpoReportActivityFileCountsCommand();
