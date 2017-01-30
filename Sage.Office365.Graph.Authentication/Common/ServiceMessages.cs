/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Authentication.Common
{
    /// <summary>
    /// Service error message constants.
    /// </summary>
    public static class ServiceMessages
    {
        /// <summary>
        /// No storage specified.
        /// </summary>
        public const string NoStore = "An authentication store is required when using the authentication provider.";

        /// <summary>
        /// Client id not specified.
        /// </summary>
        public const string NoClientId = "The client id is required to authenticate using the authentication provider.";

        /// <summary>
        /// Client secret not specified.
        /// </summary>
        public const string NoClientSecret = "The client secret is required to authenticate using the authentication provider.";

        /// <summary>
        /// Tenant id not specified.
        /// </summary>
        public const string NoTenantId = "The tenant id is required to authenticate using the authentication provider.";

        /// <summary>
        /// Redirect uri not specified.
        /// </summary>
        public const string NoRedirectUri = "The redirect uri is required to authenticate using the authentication provider.";

        /// <summary>
        /// No scopes specified.
        /// </summary>
        public const string NoScopes = "No scopes have been requested for authentication.";
        
        /// <summary>
        /// No token return from code exchange.
        /// </summary>
        public const string NoTokenReturned = "Authentication failure: no response values returned from the token authentication flow.";

        /// <summary>
        /// Failed to obtain a token for the user.
        /// </summary>
        public const string NoToken = "Authentication failure: failed to retrieve a valid authentication token.";

        /// <summary>
        /// User cancel.
        /// </summary>
        public const string UserCancel = "Authentication failure: the user cancelled authentication.";
    }
}
