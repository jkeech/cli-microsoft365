import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';


class OutlookReportMailAppUsageVersionsUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILAPPUSAGEVERSIONSUSERCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailAppUsageVersionsUserCounts';
  }

  public get description(): string {
    return 'Gets the count of unique users by Outlook desktop version.';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the count of unique users by Outlook desktop version for the last week
      m365 ${this.name} --period D7

    Gets the count of unique users by Outlook desktop version for the last week
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "mailappusageversionusercounts.txt"

    Gets the count of unique users by Outlook desktop version for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "mailappusageversionusercounts.json"
`);
  }
}

module.exports = new OutlookReportMailAppUsageVersionsUserCountsCommand();