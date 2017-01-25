/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using System;
using System.Diagnostics;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Class for maintaining authentication refresh token in testing.
    /// </summary>
    public class DebugAuthenticationStore : IAuthenticationStore
    {
        #region Private fields

        private readonly Scope _scope;
        private readonly string _clientId;
        private string _refreshToken;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="clientId">The client id to maintain the refresh token for.</param>
        public DebugAuthenticationStore(string clientId, Scope scope)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _scope = scope;

            Debug.WriteLine(string.Format("DebugAuthenticationStore({0}, {1})", _clientId, _scope));
        }

        #endregion

        #region Public properties

        /// <summary>
        /// The client id being managed by the storage class.
        /// </summary>
        public string ClientId
        {
            get { return _clientId; }
        }

        /// <summary>
        /// The refresh token being maintained for the client.
        /// </summary>
        public string RefreshToken
        {
            get { return _refreshToken; }
            set
            {
                _refreshToken = value;

                Debug.WriteLine(string.Format("RefreshToken={0}", value));
            }
        }

        /// <summary>
        /// Returns the scope for the storage.
        /// </summary>
        public Scope Scope
        {
            get { return _scope; }
        }

        #endregion
    }
}
