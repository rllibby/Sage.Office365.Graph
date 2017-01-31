/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication.Interfaces;
using System.Collections.Generic;

namespace Sage.Office365.Graph.Interfaces
{
    /// <summary>
    /// Interface definition for client implementation.
    /// </summary>
    public interface IClient
    {
        #region Public methods

        /// <summary>
        /// Gets the OneDrive based on the current principal.
        /// </summary>
        /// <returns>The instance of OneDrive for the current principal.</returns>
        IOneDrive OneDrive();

        /// <summary>
        /// Sets the user principal by identifier.
        /// </summary>
        /// <param name="principalId">The identifier of the principal to lookup.</param>
        /// <returns>True if the principal was found and set.</returns>
        void SetPrincipal(string principalId);

        /// <summary>
        /// Sign the user in to the client.
        /// </summary>
        void SignIn();

        /// <summary>
        /// Sign the user out of the client.
        /// </summary>
        void SignOut();

        /// <summary>
        /// Get the collection of users.
        /// </summary>
        /// <returns>The collection of users.</returns>
        IList<User> Users();
    
        #endregion

        #region Public properties

        /// <summary>
        /// Returns the authentication provider.
        /// </summary>
        IAuthenticationProvider2 Provider { get; }

        /// <summary>
        /// Returns the instance of the graph service client.
        /// </summary>
        GraphServiceClient ServiceClient { get; }

        /// <summary>
        /// The current user princial.
        /// </summary>
        User Principal { get; set; }

        /// <summary>
        /// Returns true if the user is signed into the graph service client.
        /// </summary>
        bool SignedIn { get; }

        #endregion
    }
}
