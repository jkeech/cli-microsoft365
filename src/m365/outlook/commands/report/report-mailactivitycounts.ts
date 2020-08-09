import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';


class OutlookReportMailActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.OUTLOOK_REPORT_MAILACTIVITYCOUNTS}`;
  }

  public get usageEndpoint(): string {
    return 'getEmailActivityCounts';
  }

  public get description(): string {
    return 'Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
      
    Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week
      m365 ${this.name} --period D7

    Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "mailactivitycounts.txt"

    Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "mailactivitycounts.json"
`);
  }
}

module.exports = new OutlookReportMailActivityCountsCommand();