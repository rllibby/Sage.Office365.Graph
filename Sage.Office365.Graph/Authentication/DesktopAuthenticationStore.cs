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
    /// Class for maintaining user isolated data for WinForm/WPF/Legacy based code.
    /// </summary>
    public sealed class DesktopAuthenticationStore : IAuthenticationStore
    {
        #region Private fields

        private static byte[] _entropy = Encoding.ASCII.GetBytes(typeof(DesktopAuthenticationStore).Name);
        private readonly Scope _scope;
        private readonly string _clientId;
        private string _refreshToken;

        #endregion

        #region Private methods

        /// <summary>
        /// Returns the isolated storage folder based on scope.
        /// </summary>
        /// <returns></returns>
        private IsolatedStorageFile GetStore()
        {
            return IsolatedStorageFile.GetStore(((_scope == Scope.User) ? IsolatedStorageScope.User : IsolatedStorageScope.Machine) | IsolatedStorageScope.Assembly, typeof(System.Security.Policy.Url), typeof(System.Security.Policy.Url));
        }

        /// <summary>
        /// Encrypts the data using data protection api.
        /// </summary>
        /// <param name="data">The data to encrypt.</param>
        /// <returns>The encrypted data on success, throws on failure.</returns>
        private byte[] Protect(byte[] data)
        {
            return ProtectedData.Protect(data, _entropy, (_scope == Scope.User) ? DataProtectionScope.CurrentUser : DataProtectionScope.LocalMachine);
        }

        /// <summary>
        /// Decrypts the data using data protection api.
        /// </summary>
        /// <param name="data">The encrypted data to decrypt.</param>
        /// <returns>The original unencrypted data.</returns>
        private byte[] Unprotect(byte[] data)
        {
            return ProtectedData.Unprotect(data, _entropy, (_scope == Scope.User) ? DataProtectionScope.CurrentUser : DataProtectionScope.LocalMachine);
        }

        /// <summary>
        /// Loads the refresh token from isolated storage.
        /// </summary>
        private void LoadRefreshToken()
        {
            byte[] persisted = null;

            try
            {
                using (var isoStore = GetStore())
                {
                    if (!isoStore.FileExists(string.Format("{0}.token", _clientId))) return;

                    using (var reader = new BinaryReader(new IsolatedStorageFileStream(string.Format("{0}.token", _clientId), FileMode.Open, isoStore)))
                    {
                        if (reader.BaseStream.Length > 0)
                        {
                            persisted = reader.ReadBytes((int)reader.BaseStream.Length);
                        }
                    }
                }
            }
            catch (Exception)
            {
                persisted = null;
            }
            finally
            {
                _refreshToken = (persisted == null) ? string.Empty : Encoding.UTF8.GetString(Unprotect(persisted));
            }
        }

        /// <summary>
        /// Saves the refresh token to isolated storage.
        /// </summary>
        /// <param name="token">The new refresh token to save.</param>
        private void SaveRefreshToken(string token)
        {
            _refreshToken = token;

            using (var isoStore = GetStore())
            {
                using (var writer = new BinaryWriter(new IsolatedStorageFileStream(string.Format("{0}.token", _clientId), FileMode.Create, isoStore)))
                {
                    if (string.IsNullOrEmpty(token))
                    {
                        writer.Write(string.Empty);

                        return;
                    }

                    var encrypted = Protect(Encoding.UTF8.GetBytes(token));

                    writer.Write(encrypted);
                }
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="clientId">The client id to maintain the refresh token for.</param>
        public DesktopAuthenticationStore(string clientId, Scope scope)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _scope = scope;

            LoadRefreshToken();
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
            set { SaveRefreshToken(value); }
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