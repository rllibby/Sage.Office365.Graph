/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Authentication.Common
{
    /// <summary>
    /// Scopes for the graph API.
    /// </summary>
    public static class GraphScopes
    {
        #region Special scopes

        public const string openid = "openid";
        public const string offline_access = "offline_access";

        #endregion

        public const string CalendarsRead = "https://graph.microsoft.com/Calendars.Read";
        public const string CalendarsReadShared = "https://graph.microsoft.com/Calendars.Read.Shared";
        public const string CalendarsReadWrite = "https://graph.microsoft.com/Calendars.ReadWrite";
        public const string CalendarsReadWriteShared = "https://graph.microsoft.com/Calendars.ReadWrite.Shared";
        public const string ContactsRead = "https://graph.microsoft.com/Contacts.Read";
        public const string ContactsReadShared = "https://graph.microsoft.com/Contacts.Read.Shared";
        public const string ContactsReadWrite = "https://graph.microsoft.com/Contacts.ReadWrite";
        public const string ContactsReadWriteShared = "https://graph.microsoft.com/Contacts.ReadWrite.Shared";
        public const string Default = "https://graph.microsoft.com/.default";
        public const string DeviceReadWriteAll = "https://graph.microsoft.com/Device.ReadWrite.All";
        public const string DirectoryAccessAsUserAll = "https://graph.microsoft.com/Directory.AccessAsUser.All";
        public const string DirectoryReadAll = "https://graph.microsoft.com/Directory.Read.All";
        public const string DirectoryReadWriteAll = "https://graph.microsoft.com/Directory.ReadWrite.All";
        public const string FilesRead = "https://graph.microsoft.com/Files.Read";
        public const string FilesReadAll = "https://graph.microsoft.com/Files.Read.All";
        public const string FilesReadSelected = "https://graph.microsoft.com/Files.Read.Selected";
        public const string FilesReadWrite = "https://graph.microsoft.com/Files.ReadWrite";
        public const string FilesReadWriteAll = "https://graph.microsoft.com/Files.ReadWrite.All";
        public const string FilesReadWriteAppFolder = "https://graph.microsoft.com/Files.ReadWrite.AppFolder";
        public const string FilesReadWriteSelected = "https://graph.microsoft.com/Files.ReadWrite.Selected";
        public const string GroupReadAll = "https://graph.microsoft.com/Group.Read.All";
        public const string GroupReadWriteAll = "https://graph.microsoft.com/Group.ReadWrite.All";
        public const string IdentityRiskEventReadAll = "https://graph.microsoft.com/IdentityRiskEvent.Read.All";
        public const string MailRead = "https://graph.microsoft.com/Mail.Read";
        public const string MailReadShared = "https://graph.microsoft.com/Mail.Read.Shared";
        public const string MailReadWrite = "https://graph.microsoft.com/Mail.ReadWrite";
        public const string MailReadWriteShared= "https://graph.microsoft.com/Mail.ReadWrite.Shared";
        public const string MailSend = "https://graph.microsoft.com/Mail.Send";
        public const string MailSendShared = "https://graph.microsoft.com/Mail.Send.Shared";
        public const string MailboxSettingsReadWrite = "https://graph.microsoft.com/MailboxSettings.ReadWrite";
        public const string MemberReadHidden = "https://graph.microsoft.com/Member.Read.Hidden";
        public const string NotesCreate = "https://graph.microsoft.com/Notes.Create";
        public const string NotesRead = "https://graph.microsoft.com/Notes.Read";
        public const string NotesReadAll = "https://graph.microsoft.com/Notes.Read.All";
        public const string NotesReadWrite = "https://graph.microsoft.com/Notes.ReadWrite";
        public const string NotesReadWriteAll = "https://graph.microsoft.com/Notes.ReadWrite.All";
        public const string NotesReadWriteCreatedByApp = "https://graph.microsoft.com/Notes.ReadWrite.CreatedByApp";
        public const string PeopleRead = "https://graph.microsoft.com/People.Read";
        public const string ReportsReadAll = "https://graph.microsoft.com/Reports.Read.All";
        public const string SitesReadAll = "https://graph.microsoft.com/Sites.Read.All";
        public const string SitesReadWriteAll = "https://graph.microsoft.com/Sites.ReadWrite.All";
        public const string TasksRead = "https://graph.microsoft.com/Tasks.Read";
        public const string TasksReadShared = "https://graph.microsoft.com/Tasks.Read.Shared";
        public const string TasksReadWrite = "https://graph.microsoft.com/Tasks.ReadWrite";
        public const string TasksReadWriteShared = "https://graph.microsoft.com/Tasks.ReadWrite.Shared";
        public const string UserRead = "https://graph.microsoft.com/User.Read";
        public const string UserReadAll = "https://graph.microsoft.com/User.Read.All";
        public const string UserReadBasicAll = "https://graph.microsoft.com/User.ReadBasic.All";
        public const string UserReadWrite = "https://graph.microsoft.com/User.ReadWrite";
        public const string UserReadWriteAll= "https://graph.microsoft.com/User.ReadWrite.All";
    }
}
