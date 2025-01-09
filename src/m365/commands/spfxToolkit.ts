import { Logger } from '../../cli/Logger';
import AnonymousCommand from '../base/AnonymousCommand';
import auth, { AuthType } from '../../Auth';
import commands from './commands';
import { AppCreationOptions, AppInfo, entraApp } from '../../utils/entraApp';
import { accessToken } from '../../utils/accessToken';

const allScopes = [
  'https://graph.windows.net/Directory.AccessAsUser.All',
  'https://management.azure.com/user_impersonation',
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/AppCatalog.ReadWrite.All',
  'https://graph.microsoft.com/AuditLog.Read.All',
  'https://graph.microsoft.com/SecurityEvents.Read.All',
  'https://graph.microsoft.com/ServiceHealth.Read.All',
  'https://graph.microsoft.com/ServiceMessage.Read.All',
  'https://graph.microsoft.com/Sites.Read.All',
  'https://graph.microsoft.com/Directory.AccessAsUser.All',
  'https://graph.microsoft.com/Directory.ReadWrite.All',
  'https://graph.microsoft.com/ChannelMember.Read.All',
  'https://graph.microsoft.com/ChannelMessage.Read.All',
  'https://graph.microsoft.com/ChannelSettings.Read.All',
  'https://graph.microsoft.com/Chat.Read',
  'https://graph.microsoft.com/OnlineMeetingArtifact.Read.All',
  'https://graph.microsoft.com/OnlineMeetings.Read',
  'https://graph.microsoft.com/OnlineMeetingTranscript.Read.All',
  'https://graph.microsoft.com/TeamMember.Read.All',
  'https://graph.microsoft.com/TeamsAppInstallation.ReadForUser',
  'https://graph.microsoft.com/TeamSettings.ReadWrite.All',
  'https://graph.microsoft.com/TeamsTab.ReadWrite.All',
  'https://manage.office.com/ActivityFeed.Read',
  'https://manage.office.com/ServiceHealth.Read',
  'https://microsoft.sharepoint-df.com/AllSites.FullControl',
  'https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All',
  'https://microsoft.sharepoint-df.com/User.ReadWrite.All'
];

class SpfxToolkitCommand extends AnonymousCommand {
  public get name(): string {
    return commands.SPFXTOOLKIT;
  }

  public get description(): string {
    return '';
  }

  public async commandAction(logger: Logger): Promise<void> {
    // setup auth
    auth.service.authType = AuthType.Browser;
    // Microsoft Azure CLI app ID
    auth.service.appId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46';
    auth.service.tenant = 'common';
    await auth.ensureAccessToken(auth.defaultResource, logger, this.debug);

    const options: AppCreationOptions = {
      allowPublicClientFlows: true,
      apisDelegated: (allScopes).join(','),
      implicitFlow: false,
      multitenant: false,
      name: 'SPFx Toolkit',
      platform: 'publicClient',
      redirectUris: 'http://localhost,https://localhost,https://login.microsoftonline.com/common/oauth2/nativeclient'
    };
    const apis = await entraApp.resolveApis({
      options,
      logger,
      verbose: this.verbose,
      debug: this.debug
    });
    const appInfo: AppInfo = await entraApp.createAppRegistration({
      options,
      apis,
      logger,
      verbose: this.verbose,
      debug: this.debug
    });
    appInfo.tenantId = accessToken.getTenantIdFromAccessToken(auth.service.accessTokens[auth.defaultResource].accessToken);
    await entraApp.grantAdminConsent({
      appInfo,
      appPermissions: entraApp.appPermissions,
      adminConsent: true,
      logger,
      debug: this.debug
    });

    await auth.clearConnectionInfo();
    auth.service.logout();

    await logger.log({
      appId: appInfo.appId,
      tenantId: appInfo.tenantId
    });
  }
}

module.exports = new SpfxToolkitCommand();