import { Logger } from '../../../../cli/Logger.js';
import { odata } from '../../../../utils/odata.js';
import { spo } from '../../../../utils/spo.js';
import SpoCommand from '../../../base/SpoCommand.js';
import commands from '../../commands.js';
import { CliRequestOptions } from "../../../../request.js";

class SpoHomeSiteListCommand extends SpoCommand {
  public get name(): string {
    return commands.HOMESITE_LIST;
  }

  public get description(): string {
    return 'Lists all home sites';
  }

  public alias(): string[] | undefined {
    return [commands.TENANT_HOMESITE_LIST];
  }

  public defaultProperties(): string[] | undefined {
    return ['Url', 'Title'];
  }

  public async commandAction(logger: Logger): Promise<void> {
    await this.showDeprecationWarning(logger, commands.TENANT_HOMESITE_LIST, commands.HOMESITE_LIST);

    try {
      const spoAdminUrl: string = await spo.getSpoAdminUrl(logger, this.verbose);
      const requestOptions: CliRequestOptions = {
        url: `${spoAdminUrl}/_api/SPO.Tenant/GetTargetedSitesDetails`,
        headers: {
          accept: 'application/json;odata=nometadata'
        },
        responseType: 'json'
      };
      if (this.verbose) {
        await logger.logToStderr(`Retrieving all home sites...`);
      }
      const res = await odata.getAllItems(requestOptions);
      await logger.log(res);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new SpoHomeSiteListCommand();