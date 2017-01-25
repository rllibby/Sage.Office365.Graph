/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using System;
using System.Security.Cryptography;
using System.IO;
using System.IO.IsolatedStorage;
using System.Text;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Protection level scope for the store.
    /// </summary>
    public enum Scope
    {
        /// <summary>
        /// The store is scoped to the user.
        /// </summary>
        User,

        /// <summary>
        /// The store is scoped to the system.
        /// </summary>
        System
    }

    /// <summary>
    /// Class for maintaining user isolated data.
    /// </summary>
    public sealed class AuthenticationStore : IDisposable
    {
        #region Private fields

        private static byte[] _entropy = Encoding.ASCII.GetBytes(typeof(AuthenticationStore).Name);
        private IsolatedStorageFile _isoStore;
        private Scope _scope;
        private string _refreshToken;
        private string _clientId;
        private bool _disposed;

        #endregion

        #region Private methods

        /// <summary>
        /// Encrypts the data using data protection api.
        /// </summary>
        /// <param name="data">The data to encrypt.</param>
        /// <returns>The encrypted data on success, throws on failure.</returns>
        public byte[] Protect(byte[] data)
        {
            return ProtectedData.Protect(data, _entropy, (_scope == Scope.User) ? DataProtectionScope.CurrentUser : DataProtectionScope.LocalMachine);
        }

        /// <summary>
        /// Decrypts the data using data protection api.
        /// </summary>
        /// <param name="data">The encrypted data to decrypt.</param>
        /// <returns>The original unencrypted data.</returns>
        public byte[] Unprotect(byte[] data)
        {
            return ProtectedData.Unprotect(data, _entropy, (_scope == Scope.User) ? DataProtectionScope.CurrentUser : DataProtectionScope.LocalMachine);
        }

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
        public AuthenticationStore(string clientId, )
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