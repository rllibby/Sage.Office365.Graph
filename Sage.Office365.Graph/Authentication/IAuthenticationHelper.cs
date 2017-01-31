/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Interface definition for the authentication helper (client/application) implementations.
    /// </summary>
    public interface IAuthenticationHelper
    {
        /// <summary>
        /// The Microsoft Graph service client.
        /// </summary>
        GraphServiceClient Client { get; }

        /// <summary>
        /// The client id for the authentication.
        /// </summary>
        string ClientId { get; }

        /// <summary>
        /// Returns true if the instance is authenticated.
        /// </summary>
        bool SignedIn { get; }

        /// <summary>
        /// Attempts to authenticate the instance with the graph service.
        /// </summary>
        void SignIn();

        /// <summary>
        /// Tears down the authentication with the graph service.
        /// </summary>
        void SignOut();
    }
}
