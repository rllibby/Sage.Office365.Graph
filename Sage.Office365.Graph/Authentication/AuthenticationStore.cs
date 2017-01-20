/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using System;
using System.Diagnostics;
using System.IO;
using System.IO.IsolatedStorage;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Class for maintaining user isolated data.
    /// </summary>
    public sealed class AuthenticationStore : IDisposable
    {
        #region Private fields

        private IsolatedStorageFile _isoStore;
        private string _refreshToken;
        private string _clientId;
        private bool _disposed;

        #endregion

        #region Private methods

        /// <summary>
        /// Loads the refresh token from isolated storage.
        /// </summary>
        private void LoadRefreshToken()
        {
            var persisted = string.Empty;

            try
            {
                if (!_isoStore.FileExists(string.Format("{0}.token", _clientId))) return;

                using (var reader = new StreamReader(new IsolatedStorageFileStream(string.Format("{0}.token", _clientId), FileMode.Open, _isoStore)))
                {
                    persisted = reader.ReadToEnd();
                }
            }
            finally
            {
                _refreshToken = persisted;
            }
        }

        /// <summary>
        /// Saves the refresh token to isolated storage.
        /// </summary>
        /// <param name="token">The new refresh token to save.</param>
        private void SaveRefreshToken(string token)
        {
            _refreshToken = token;

            using (var writer = new StreamWriter(new IsolatedStorageFileStream(string.Format("{0}.token", _clientId), FileMode.Create, _isoStore)))
            {
                writer.Write(_refreshToken);
            }
        }

        /// <summary>
        /// Resource cleanup.
        /// </summary>
        /// <param name="disposing">True if being disposed.</param>
        private void Dispose(bool disposing)
        {
            if (!disposing || _disposed) return;

            try
            {
                if (_isoStore != null)
                {
                    _isoStore.Dispose();
                    _isoStore = null;
                }
            }
            finally
            {
                _disposed = true;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="clientId">The client id to maintain the refresh token for.</param>
        public AuthenticationStore(string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _isoStore = IsolatedStorageFile.GetStore(IsolatedStorageScope.User | IsolatedStorageScope.Assembly, typeof(System.Security.Policy.Url), typeof(System.Security.Policy.Url));

            LoadRefreshToken();
        }

        #endregion

        #region Public methods

        /// <summary>
        ///  Resource cleanup.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Private properties

        /// <summary>
        /// The refresh token being maintained for the client.
        /// </summary>
        public string RefreshToken
        {
            get { return _refreshToken; }
            set { SaveRefreshToken(value); }
        }

        #endregion
    }
}