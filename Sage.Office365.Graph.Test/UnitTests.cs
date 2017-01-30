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
using System.Diagnostics;

namespace Sage.Office365.Graph.Test
{
    [TestClass]
    public class UnitTests
    {
        #region Public constants

        public const string ClientId = "b85030a0-4aaa-4777-a80c-5553f949f745";
        public const string ClientSecret = "gEREtebjT3HedPakTTP2Rrc";
        public const string TenantId = "b366b438-32ef-4f31-81f9-0c2214ec7d87";
        public const string AdminRedirect = "http://localhost/sagepaperless";

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
    }
}
