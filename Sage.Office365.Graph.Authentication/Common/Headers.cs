/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Authentication.Common
{
    /// <summary>
    /// Authentication header constants.
    /// </summary>
    public static class Headers
    {
        public const string Bearer = "Bearer";
        public const string SdkVersionHeaderName = "X-ClientService-ClientTag";
        public const string FormUrlEncodedContentType = "application/x-www-form-urlencoded";
        public const string SdkVersionHeaderValue = "SDK-Version=CSharp-v{0}";
        public const string ThrowSiteHeaderName = "X-ThrowSite";
    }
}
