/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Sage.Office365.Graph.Authentication;
using Sage.Office365.Graph.Authentication.Common;
using Sage.Office365.Graph.Authentication.Interfaces;
using Sage.Office365.Graph.Authentication.Storage;
using System;
using System.Diagnostics;

namespace Sage.Office365.Graph.Test
{
    [TestClass]
    public class UnitTests
    {
        #region Public constants

        public const string ClientId = "7a2d47f6-83d7-49da-b17d-6afd5a072e4e";
        public const string ClientSecret = "56FfAg8oDAKRxhO6UnQgwFt";
        public const string TenantId = "d869ff02-e7b3-436c-8e59-437c755b965c";
        public const string AdminRedirect = "http://localhost/sage100USdev";

        #endregion

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderUser()
        {
            var provider = new AuthenticationProvider(ClientId);

            Assert.IsNotNull(provider);
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderApp()
        {
            var provider = new AuthenticationProvider(ClientId, ClientSecret, TenantId);

            Assert.IsNotNull(provider);
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderAuth()
        {
            var provider = new AuthenticationProvider(ClientId);

            Assert.IsNotNull(provider);

            var task = provider.AuthenticateAsync();

            task.Wait();

            Assert.IsTrue(provider.Authenticated);
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderAuthStore()
        {
            var store = new DesktopAuthenticationStore(ClientId, Scope.User);
            var provider = new AuthenticationProvider(ClientId, store);

            Assert.IsNotNull(provider);

            var task = provider.AuthenticateAsync();

            task.Wait();

            Assert.IsTrue(provider.Authenticated);

            provider = new AuthenticationProvider(ClientId, store);

            task = provider.AuthenticateAsync();
            task.Wait();
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderLogout()
        {
            var store = new DesktopAuthenticationStore(ClientId, Scope.User);
            var provider = new AuthenticationProvider(ClientId, store);

            Assert.IsNotNull(provider);

            var task = provider.AuthenticateAsync();

            task.Wait();

            provider.Logout();
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderAuthApp()
        {
            var provider = new AuthenticationProvider(ClientId, ClientSecret, TenantId);

            Assert.IsNotNull(provider);

            var task = provider.AuthenticateAsync();

            task.Wait();

            Assert.IsTrue(provider.Authenticated);
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderAdminConsent()
        {
            var provider = new AuthenticationProvider(ClientId, ClientSecret, TenantId);

            Assert.IsNotNull(provider);
            Assert.IsTrue(provider.GetAdminConsent(AdminRedirect));
        }

        [TestMethod]
        [TestCategory("Provider")]
        [ExpectedException(typeof(ServiceException))]
        public void Test_ProviderAdminConsentBadRedirect()
        {
            var provider = new AuthenticationProvider(ClientId, ClientSecret, TenantId);

            Assert.IsNotNull(provider);
            Assert.IsTrue(provider.GetAdminConsent("http://localhost"));
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderAppProperties()
        {
            var provider = new AuthenticationProvider(ClientId, ClientSecret, TenantId);
            var task = provider.AuthenticateAsync();

            task.Wait();

            Assert.IsTrue(provider.AppBased);
            Assert.IsTrue(provider.Authenticated);
            Assert.IsFalse(string.IsNullOrEmpty(provider.AccessToken));
            Assert.IsTrue(provider.ClientId.Equals(ClientId));
            Assert.IsNull(provider.RefreshToken);
            Assert.IsTrue(provider.Scopes.Contains(GraphScopes.Default));
            Assert.IsTrue(provider.TenantId.Equals(TenantId));

            provider.Logout();
        }

        [TestMethod]
        [TestCategory("Provider")]
        public void Test_ProviderUserProperties()
        {
            var provider = new AuthenticationProvider(ClientId);

            provider.Scopes.Add(GraphScopes.UserReadWriteAll);
            provider.Scopes.Add(GraphScopes.FilesReadWriteAll);
            provider.Scopes.Add(GraphScopes.GroupReadWriteAll);
            provider.Scopes.Add(GraphScopes.DirectoryAccessAsUserAll);

            var task = provider.AuthenticateAsync();

            task.Wait();

            Assert.IsFalse(provider.AppBased);
            Assert.IsTrue(provider.Authenticated);
            Assert.IsFalse(string.IsNullOrEmpty(provider.AccessToken));
            Assert.IsTrue(provider.ClientId.Equals(ClientId));
            Assert.IsNotNull(provider.RefreshToken);
            Assert.IsNull(provider.TenantId);

            provider.Logout();
        }

        [TestMethod]
        [TestCategory("Client")]
        public void Test_ClientRoot()
        {
            var userClient = new Client(ClientId);

            userClient.Provider.Scopes.Add(GraphScopes.UserReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.FilesReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.GroupReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.DirectoryAccessAsUserAll);

            userClient.SignIn();

            var drive = userClient.OneDrive();
            var root = drive.Root;

        }

        [TestMethod]
        [TestCategory("Client")]
        public void Test_ClientExample()
        {
            var userClient = new Client(ClientId);

            userClient.Provider.Scopes.Add(GraphScopes.UserReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.FilesReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.GroupReadWriteAll);
            userClient.Provider.Scopes.Add(GraphScopes.DirectoryAccessAsUserAll);

            userClient.SignIn();

            var appClient = new Client(ClientId, ClientSecret, TenantId);

            appClient.Provider.GetAdminConsent("http://localhost/sage100USdev");

            appClient.SignIn();

            Debug.WriteLine(string.Format("User Based Auth: {0}", userClient.Principal?.DisplayName));
            Debug.WriteLine(string.Format("App Based Auth: {0}", appClient.Principal?.DisplayName));

            var users = appClient.Users();

            foreach (var user in users) Debug.WriteLine(user.DisplayName);

            appClient.SetPrincipal("russell@sage100USdev.onmicrosoft.com");

            Debug.WriteLine(string.Format("App Based Auth: {0}", appClient.Principal?.DisplayName));

            var client = appClient; // userClient;

            var drive = client.OneDrive();
            var folders = drive.GetChildren();
            var firstFolder = folders[0];

            var temp = drive.CreateFolder(firstFolder, Guid.NewGuid().ToString());

            var file1 = drive.UploadFile(null, @"c:\temp\ESTemplates.pdf");
            var file2 = drive.UploadFile(temp, @"c:\temp\ITTemplates.pdf");

            var subFolders = drive.GetChildren(firstFolder);

            var pdf = drive.GetItem(file1);

            drive.DownloadFile(pdf, "c:\\temp", true);
            drive.DeleteItem(temp);
        }
    }
}
