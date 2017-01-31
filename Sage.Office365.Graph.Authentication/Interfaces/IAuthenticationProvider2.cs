/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Sage.Office365.Graph.Authentication.Interfaces
{
    /// <summary>
    /// Extended version of the authentication provider.
    /// </summary>
    public interface IAuthenticationProvider2 : IAuthenticationProvider
    {
        #region Public methods

        /// <summary>
        /// Attempts to authenticate either the app or the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        Task AuthenticateAsync();

        /// <summary>
        /// Gets the administration consent to allow application to access resources.
        /// </summary>
        /// <param name="redirectUri">The redirect uri for the oAuth request.</param>
        /// <returns>True if consent was given, otherwise false.</returns>
        bool GetAdminConsent(string redirectUri);

        /// <summary>
        /// Logs the user or app out of the tenant.
        /// </summary>
        void Logout();

        #endregion

        #region Public properties

        /// <summary>
        /// Returns true if configured for application authentication.
        /// </summary>
        bool AppBased { get; }

        /// <summary>
        /// Returns true if the provider is authenticated. 
        /// </summary>
        bool Authenticated { get; }

        /// <summary>
        /// Returns the access token.
        /// </summary>
        string AccessToken { get; }

        /// <summary>
        /// Returns the client id.
        /// </summary>
        string ClientId { get; }

        /// <summary>
        /// The refresh token.
        /// </summary>
        string RefreshToken { get; set; }

        /// <summary>
        /// The collection of scopes to request when authenticating.
        /// </summary>
        IList<string> Scopes { get; }

        /// <summary>
        /// Returns the tenant id.
        /// </summary>
        string TenantId { get; }

        #endregion
    }
}
