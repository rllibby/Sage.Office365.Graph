/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace Sage.Office365.Graph.Test
{
    [TestClass]
    public class UnitTests
    {
        #region Public constants

        public const string ValidClientId = "2585caa5-8695-4945-9da7-6d6a73c5b172";
        public const string InvalidClientId = "2995caa5-8695-4945-9da7-6d6a73c5b172";

        #endregion

        [TestMethod]
        [ExpectedException(typeof(ApplicationException))]
        public void Test_InvalidClientId()
        {
            var client = Client.Create(InvalidClientId);
        }

        [TestMethod]
        public void Test_ValidClientId()
        {
            var client = Client.Create(ValidClientId);
        }

        [TestMethod]
        public void Test_SignOut()
        {
            var client = Client.Create(ValidClientId);

            client.SignOut();
        }

        [TestMethod]
        public void Test_Me()
        {
            var client = Client.Create(ValidClientId);
            var me = client.Me();

            Assert.IsFalse(string.IsNullOrEmpty(me.DisplayName));
            Assert.IsFalse(string.IsNullOrEmpty(me.UserPrincipalName));
        }

        [TestMethod]
        public void Test_RootItems()
        {
            var client = Client.Create(ValidClientId);
            var items = client.GetRootItems();

            Assert.IsFalse(items.Count == 0);
        }
    }
}
