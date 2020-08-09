import commands from '../../commands';
import DateAndPeriodBasedReport from '../../../base/DateAndPeriodBasedReport';


class TenantReportActiveUserDetailCommand extends DateAndPeriodBasedReport {
  public get name(): string {
    return commands.TENANT_REPORT_ACTIVEUSERDETAIL;
  }

  public get description(): string {
    return 'Gets details about Microsoft 365 active users';
  }

  public get usageEndpoint(): string {
    return 'getOffice365ActiveUserDetail';
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    As this report is only available for the past 28 days, date parameter value
    should be a date from that range.

  Examples:

    Gets details about Microsoft 365 active users for the last week
      m365 ${this.name} --period D7

    Gets details about Microsoft 365 active users for May 1, 2019
      m365 ${this.name} --date 2019-05-01

    Gets details about Microsoft 365 active users for the last week 
    and exports the report data in the specified path in text format
      m365 ${this.name} --period D7 --output text > "activeuserdetail.txt"

    Gets details about Microsoft 365 active users for the last week
    and exports the report data in the specified path in json format
      m365 ${this.name} --period D7 --output json > "activeuserdetail.json"
  `);
  }
}

module.exports = new TenantReportActiveUserDetailCommand();