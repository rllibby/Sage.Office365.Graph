/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Authentication.Interfaces
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
    /// Interface definition for authentication storage used to handle the refresh token.
    /// </summary>
    public interface IAuthenticationStore
    {
        /// <summary>
        /// The client id for the storage.
        /// </summary>
        string ClientId { get; }

        /// <summary>
        /// The refresh token.
        /// </summary>
        string RefreshToken { get; set; }

        /// <summary>
        /// The scope for the storage; either machine or user based.
        /// </summary>
        Scope Scope { get; }
    }
}
