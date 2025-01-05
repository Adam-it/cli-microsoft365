import { app } from './utils/app.js';

export default {
  allScopes: [
    'https://graph.windows.net/Directory.AccessAsUser.All',
    'https://management.azure.com/user_impersonation',
    'https://admin.services.crm.dynamics.com/user_impersonation',
    'https://graph.microsoft.com/AppCatalog.ReadWrite.All',
    'https://graph.microsoft.com/AuditLog.Read.All',
    'https://graph.microsoft.com/Bookings.Read.All',
    'https://graph.microsoft.com/Calendars.Read',
    'https://graph.microsoft.com/ChannelMember.ReadWrite.All',
    'https://graph.microsoft.com/ChannelMessage.Read.All',
    'https://graph.microsoft.com/ChannelMessage.ReadWrite',
    'https://graph.microsoft.com/ChannelMessage.Send',
    'https://graph.microsoft.com/ChannelSettings.ReadWrite.All',
    'https://graph.microsoft.com/Chat.ReadWrite',
    'https://graph.microsoft.com/Community.ReadWrite.All',
    'https://graph.microsoft.com/Directory.AccessAsUser.All',
    'https://graph.microsoft.com/Directory.ReadWrite.All',
    'https://graph.microsoft.com/ExternalConnection.ReadWrite.All',
    'https://graph.microsoft.com/ExternalItem.ReadWrite.All',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/IdentityProvider.ReadWrite.All',
    'https://graph.microsoft.com/InformationProtectionPolicy.Read',
    'https://graph.microsoft.com/Mail.Read.Shared',
    'https://graph.microsoft.com/Mail.ReadWrite',
    'https://graph.microsoft.com/Mail.Send',
    'https://graph.microsoft.com/MailboxSettings.ReadWrite',
    'https://graph.microsoft.com/Notes.ReadWrite.All',
    'https://graph.microsoft.com/OnlineMeetingArtifact.Read.All',
    'https://graph.microsoft.com/OnlineMeetings.ReadWrite',
    'https://graph.microsoft.com/OnlineMeetingTranscript.Read.All',
    'https://graph.microsoft.com/PeopleSettings.ReadWrite.All',
    'https://graph.microsoft.com/Place.Read.All',
    'https://graph.microsoft.com/Policy.Read.All',
    'https://graph.microsoft.com/RecordsManagement.ReadWrite.All',
    'https://graph.microsoft.com/Reports.ReadWrite.All',
    'https://graph.microsoft.com/RoleAssignmentSchedule.ReadWrite.Directory',
    'https://graph.microsoft.com/RoleEligibilitySchedule.Read.Directory',
    'https://graph.microsoft.com/SecurityEvents.Read.All',
    'https://graph.microsoft.com/ServiceHealth.Read.All',
    'https://graph.microsoft.com/ServiceMessage.Read.All',
    'https://graph.microsoft.com/ServiceMessageViewpoint.Write',
    'https://graph.microsoft.com/Sites.Read.All',
    'https://graph.microsoft.com/Tasks.ReadWrite',
    'https://graph.microsoft.com/Team.Create',
    'https://graph.microsoft.com/TeamMember.ReadWrite.All',
    'https://graph.microsoft.com/TeamsAppInstallation.ReadWriteForUser',
    'https://graph.microsoft.com/TeamSettings.ReadWrite.All',
    'https://graph.microsoft.com/TeamsTab.ReadWrite.All',
    'https://graph.microsoft.com/User.Invite.All',
    'https://manage.office.com/ActivityFeed.Read',
    'https://manage.office.com/ServiceHealth.Read',
    'https://analysis.windows.net/powerbi/api/Dataset.Read.All',
    'https://api.powerapps.com//User',
    'https://microsoft.sharepoint-df.com/AllSites.FullControl',
    'https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All',
    'https://microsoft.sharepoint-df.com/User.ReadWrite.All'
  ],
  applicationName: `CLI for Microsoft 365 v${app.packageJson().version}`,
  delimiter: 'm365\$',
  configstoreName: 'cli-m365-config',
  minimalScopes: [
    'https://graph.microsoft.com/User.Read'
  ]
};