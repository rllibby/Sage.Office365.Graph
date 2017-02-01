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
    /// Helper class for creating an authenticated Graph client based on a user.
    /// </summary>
    public class AuthenticationHelper : IAuthentication
    {
        #region Private constants

        private readonly IAuthenticationStore _store;
        private GraphServiceClient _graphClient;
        private readonly object _lock = new object();
        private readonly string _clientId;
        private readonly static string[] _scopes =   {
                                                "offline_access",
                                                "https://graph.microsoft.com/User.ReadWrite.All",
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
        /// <returns>The async task.</returns>
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
        /// <param name="store">The authentication store for handling the refresh token.</param>
        /// <param name="clientId">The client id to create the helper for.</param>
        public AuthenticationHelper(IAuthenticationStore store, string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _store = (store == null) ? new DesktopAuthenticationStore(clientId, Scope.System) : store; 
        }

        #endregion

        #region Public methods

        /// <summary>
        ///  Attempts to authenticate the instance with the graph service.
        /// </summary>
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
        /// Tears down the authentication with the graph service.
        /// </summary>
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
        /// The client id for the authentication.
        /// </summary>
        public string ClientId
        {
            get { return _clientId; }
        }

        /// <summary>
        /// Returns true if the instance is authenticated.
        /// </summary>
        public bool SignedIn
        {
            get { return (_graphClient != null); }
        }

        #endregion
    }

    /// <summary>
    /// Helper class for creating an authenticated Graph client based on an application.
    /// </summary>
    public class AppAuthenticationHelper : IAuthenticationHelper
    {
        #region Private constants

        private GraphServiceClient _graphClient;
        private readonly object _lock = new object();
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;
        private readonly static string[] _scopes = { "https://graph.microsoft.com/.default" };

        #endregion

        #region Private methods

        /// <summary>
        /// Aquires a new authenticated instance of the GraphClient.
        /// </summary>
        /// <returns>The async task</returns>
        private async Task AquireGraphClient(string clientId, string clientSecret, string tenantId)
        {
            var authenticationProvider = new OAuth2AppAuthenticationProvider(clientId, clientSecret, tenantId);

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
        /// <param name="clientId">The client id to create the client for.</param>
        /// <param name="clientSecret">The client secret used for application authentication.</param>
        /// <param name="tenantId">The tenant id used for application authentication.</param>
        public AppAuthenticationHelper(string clientId, string clientSecret, string tenantId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");
            if (string.IsNullOrEmpty(clientSecret)) throw new ArgumentNullException("clientSecret");
            if (string.IsNullOrEmpty(tenantId)) throw new ArgumentNullException("tenantId");

            _clientId = clientId;
            _clientSecret = clientSecret;
            _tenantId = tenantId;
        }

        #endregion

        #region Public methods

        /// <summary>
        ///  Attempts to authenticate the instance with the graph service.
        /// </summary>
        public void SignIn()
        {
            if (_graphClient != null) return;

            lock (_lock)
            {
                try
                {
                    if (_graphClient != null) return;

                    var task = AquireGraphClient(_clientId, _clientSecret, _tenantId);

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
        /// Tears down the authentication with the graph service.
        /// </summary>
        public void SignOut()
        {
            _graphClient = null;
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
        /// The client id for the authentication.
        /// </summary>
        public string ClientId
        {
            get { return _clientId; }
        }

        /// <summary>
        /// Returns true if the instance is authenticated.
        /// </summary>
        public bool SignedIn
        {
            get { return (_graphClient != null); }
        }

        #endregion
    }
}
