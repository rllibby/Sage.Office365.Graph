/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Extensions;
using System;
using System.Threading.Tasks;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Helper class for creating an authenticated Graph client.
    /// </summary>
    public class AuthenticationHelper
    {
        #region Private constants

        private readonly IAuthenticationStore _store;
        private GraphServiceClient _graphClient;
        private readonly object _lock = new object();
        private readonly string _clientId;
        private readonly static string[] _scopes =   {
                                                "offline_access",
                                                "https://graph.microsoft.com/Files.ReadWrite",
                                                "https://graph.microsoft.com/Group.ReadWrite.All",
                                                "https://graph.microsoft.com/Directory.ReadWrite.All",
                                                "https://graph.microsoft.com/Directory.AccessAsUser.All",
                                            };

        #endregion

        #region Private methods

        /// <summary>
        /// Aquires a new authenticated instance of the GraphClient.
        /// </summary>
        /// <returns></returns>
        private async Task AquireGraphClient(string clientId)
        {
            var authenticationProvider = new OAuth2AuthenticationProvider(_store, clientId, Common.Constants.RedirectUri, _scopes);

            try
            {
                await authenticationProvider.AuthenticateAsync();

                _graphClient = new GraphServiceClient(authenticationProvider);
            }
            catch (Exception)
            {
                _graphClient = null;

                throw;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="clientId">The client id to create the helper for.</param>
        public AuthenticationHelper(IAuthenticationStore store, string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _store = (store == null) ? new AuthenticationStore(clientId, Scope.System) : store; 
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Attempts to sign the user in and obtain the Microsoft Graph access token.
        /// </summary>
        /// <returns>The async task which can be awaited.</returns>
        public void SignIn()
        {
            if (_graphClient != null) return;

            lock (_lock)
            { 
                try
                {
                    if (_graphClient != null) return;

                    var task = AquireGraphClient(_clientId);

                    task.WaitWithPumping();

                    if (task.IsFaulted) throw task.Exception;
                }
                catch
                {
                    _graphClient = null;

                    throw;
                }
            }
        }

        /// <summary>
        /// Attempts to sign the user out of Microsoft Graph.
        /// </summary>
        /// <returns>The async task which can be awaited.</returns>
        public void SignOut()
        {
            try
            {
                _store.RefreshToken = string.Empty;
            }
            finally
            {
                _graphClient = null;
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// The Microsoft Graph service client.
        /// </summary>
        public GraphServiceClient Client
        {
            get { return _graphClient; }
            set { _graphClient = value; }
        }

        /// <summary>
        /// Returns true if the user is signed in, otherwise false.
        /// </summary>
        public bool SignedIn
        {
            get { return (_graphClient != null); }
        }

        #endregion
    }
}
