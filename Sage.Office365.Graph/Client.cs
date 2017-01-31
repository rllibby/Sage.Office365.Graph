/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication;
using Sage.Office365.Graph.Authentication.Interfaces;
using Sage.Office365.Graph.Authentication.Storage;
using Sage.Office365.Graph.Extensions;
using Sage.Office365.Graph.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;

namespace Sage.Office365.Graph
{
    /// <summary>
    /// Client class for Sage integration with Office 365 OneDrive
    /// </summary>
    public class Client : SynchronousHandler, IClient
    {
        #region Private fields

        private IAuthenticationProvider2 _provider;
        private GraphServiceClient _serviceClient;
        private User _principal;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for user based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        public Client(string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _provider = new AuthenticationProvider(clientId, new DesktopAuthenticationStore(clientId, Scope.System));
        }

        /// <summary>
        /// Constructor for app based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="clientSecret">The client secret which will be used to perform application authentication.</param>
        /// <param name="tenantId">The Office 365 tenant id to perform application authentication.</param>
        public Client(string clientId, string clientSecret, string tenantId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");
            if (string.IsNullOrEmpty(clientSecret)) throw new ArgumentNullException("clientSecret");
            if (string.IsNullOrEmpty(tenantId)) throw new ArgumentNullException("tenantId");

            _provider = new AuthenticationProvider(clientId, clientSecret, tenantId);
       }

        #endregion

        #region Public methods

        /// <summary>
        /// Gets the OneDrive based on the current principal.
        /// </summary>
        /// <returns>The instance of OneDrive for the current principal.</returns>
        public IOneDrive OneDrive()
        {
            SignIn();

            if (_principal == null) ThrowException(new ServiceException(new Error { Code = GraphErrorCode.InvalidRequest.ToString(), Message = "No user principal set for the client." }));

            return new OneDrive(this, _principal);
        }

        /// <summary>
        /// Sets the user principal by identifier.
        /// </summary>
        /// <param name="principalId">The identifier of the principal to lookup.</param>
        /// <returns>True if the principal was found and set.</returns>
        public void SetPrincipal(string principalId)
        {
            if (string.IsNullOrEmpty(principalId)) throw new ArgumentNullException("principalId");

            SignIn();

            _principal = ExecuteTask(_serviceClient.Users[principalId].Request().GetAsync());
        }

        /// <summary>
        /// Sign the user in to the client.
        /// </summary>
        public void SignIn()
        {
            if (SignedIn) return;

            try
            {
                if (!_provider.Authenticated) ExecuteTask(_provider.AuthenticateAsync());

                _serviceClient = new GraphServiceClient(_provider);

                if (!_provider.AppBased) _principal = ExecuteTask(_serviceClient.Me.Request().GetAsync());
            }
            catch (Exception)
            {
                _serviceClient = null;

                throw;
            }
        }

        /// <summary>
        /// Sign the user out of the client.
        /// </summary>
        public void SignOut()
        {
            try
            {
                if (_provider.Authenticated) _provider.Logout();
            }
            finally
            {
                _principal = null;
                _serviceClient = null;
            }
        }

        /// <summary>
        /// Get the collection of users.
        /// </summary>
        /// <returns>The collection of users.</returns>
        public IList<User> Users()
        {
            SignIn();

            var result = new List<User>();
            var userPage = ExecuteTask(_serviceClient.Users.Request().GetAsync());

            while (userPage != null)
            {
                foreach (var user in userPage.CurrentPage)
                {
                    result.Add(user);
                }

                userPage = (userPage.NextPageRequest == null) ? null : ExecuteTask(userPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns the authentication provider.
        /// </summary>
        public IAuthenticationProvider2 Provider
        {
            get { return _provider; }
        }

        /// <summary>
        /// Returns the instance of the graph service client.
        /// </summary>
        public GraphServiceClient ServiceClient
        {
            get { return _serviceClient; }
        }

        /// <summary>
        /// The current user princial.
        /// </summary>
        public User Principal
        {
            get { return _principal; }
            set { _principal = value; }
        }

        /// <summary>
        /// Returns true if the user is signed into the graph service client.
        /// </summary>
        public bool SignedIn
        {
            get { return ((_serviceClient != null) && _provider.Authenticated); }
        }

        #endregion
    }
}
