import { Logger } from '../../../../cli/Logger';
import GlobalOptions from '../../../../GlobalOptions';
import { odata } from '../../../../utils/odata';
import { spo } from '../../../../utils/spo';
import SpoCommand from '../../../base/SpoCommand';
import commands from '../../commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  excludeDeletedSites: boolean;
}

class SpoSiteAppCatalogListCommand extends SpoCommand {

  public get name(): string {
    return commands.SITE_APPCATALOG_LIST;
  }

  public get description(): string {
    return 'List all site collection app catalogs within the tenant';
  }

  constructor() {
    super();

    this.#initOptions();
  }

  #initOptions(): void {
    this.options.unshift(
      {
        option: '--excludeDeletedSites'
      }
    );
  }

  public defaultProperties(): string[] | undefined {
    return ['AbsoluteUrl', 'SiteID'];
  }

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        logger.logToStderr('Retrieving site collection app catalogs...');
      }

      const spoUrl: string = await spo.getSpoUrl(logger, this.debug);
      let appCatalogs = await odata.getAllItems<any>(`${spoUrl}/_api/Web/TenantAppCatalog/SiteCollectionAppCatalogsSites`);

      if (args.options.excludeDeletedSites) {
        if (this.verbose) {
          logger.logToStderr('Excluding inaccessible sites from the results...');
        }

        const activeAppCatalogs = [];
        for (const appCatalog of appCatalogs) {
          try {
            await spo.getWeb(appCatalog.AbsoluteUrl, logger, this.verbose);
            activeAppCatalogs.push(appCatalog);
          }
          catch (error: any) {
            if (this.debug) {
              logger.logToStderr(error);
            }

            if (error.response.status === 404 || error.response.status === 403) {
              if (this.verbose) {
                logger.logToStderr(`Site at '${appCatalog.AbsoluteUrl}' is inaccessible. Excluding from results...`);
              }
              continue;
            }

            throw error;
          }
        }

        appCatalogs = activeAppCatalogs;
      }

      await logger.log(appCatalogs);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

module.exports = new SpoSiteAppCatalogListCommand();