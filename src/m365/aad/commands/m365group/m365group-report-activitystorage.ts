import PeriodBasedReport from '../../../base/PeriodBasedReport';
import commands from '../../commands';

class M365GroupReportActivityStorageCommand extends PeriodBasedReport {
  public get name(): string {
    return commands.M365GROUP_REPORT_ACTIVITYSTORAGE;
  }

  public get description(): string {
    return 'Get the total storage used across all group mailboxes and group sites';
  }

  public get usageEndpoint(): string {
    return 'getOffice365GroupsActivityStorage';
  }
}

module.exports = new M365GroupReportActivityStorageCommand();