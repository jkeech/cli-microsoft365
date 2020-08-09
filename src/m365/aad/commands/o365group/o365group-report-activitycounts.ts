import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';


class O365GroupReportActivityCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return `${commands.O365GROUP_REPORT_ACTIVITYCOUNTS}`;
  }

  public get description(): string {
    return 'Get the number of group activities across group workloads';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:

    Get the number of group activities across group workloads for the last week
      m365 ${this.name} --period D7

    Get the number of group activities across group workloads for the last week
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "o365groupactivitycounts.txt"

    Get the number of group activities across group workloads for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "o365groupactivitycounts.json"
`);
  }
}

module.exports = new O365GroupReportActivityCountsCommand();