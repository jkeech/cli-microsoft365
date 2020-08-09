import commands from '../../commands';
import PeriodBasedReport from '../../../base/PeriodBasedReport';


class TenantReportServicesUserCountsCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.TENANT_REPORT_SERVICESUSERCOUNTS;
  }

  public get description(): string {
    return 'Gets the count of users by activity type and service.';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ServicesUserCounts';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples: 

    Gets the count of users by activity type and service for the last week
      m365 ${this.name} --period D7

    Gets the count of users by activity type and service for the last week
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "servicesusercount.txt"

    Gets the count of users by activity type and service for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "servicesusercount.json"
`);
  }
}

module.exports = new TenantReportServicesUserCountsCommand();