/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Authentication.Common
{
    /// <summary>
    /// Authentication endpoints.
    /// </summary>
    public static class Endpoints
    {
        /// <summary>
        /// Administrator consent endpoint.
        /// </summary>
        public const string AdminConsentServiceUrl = "https://login.microsoftonline.com/{0}/adminconsent";

        /// <summary>
        /// User authentication endpoint.
        /// </summary>
        public const string AuthenticationServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

        /// <summary>
        /// User authentication token endpoint.
        /// </summary>
        public const string AuthenticationTokenServiceUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        /// <summary>
        /// The redirect uri for logout operations.
        /// </summary>
        public const string LogoutRedirectUrl = "http://localhost";

        /// <summary>
        /// User logout endpoint.
        /// </summary>
        public const string LogoutServiceUrl = "https://login.microsoftonline.com/common/oauth2/logout";

        /// <summary>
        /// The redirect uri for native applications.
        /// </summary>
        public const string NativeRedirectUrl = "urn:ietf:wg:oauth:2.0:oob";

        /// <summary>
        /// Application authentication token endpoint.
        /// </summary>
        public const string TokenServiceUrl = "https://login.microsoftonline.com/{0}/oauth2/v2.0/token";
    }
}
