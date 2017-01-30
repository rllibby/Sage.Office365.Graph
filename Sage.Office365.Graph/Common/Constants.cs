/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

namespace Sage.Office365.Graph.Common
{
    /// <summary>
    /// Static class for maintaining the assembly constants.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Root drive reference for Office 365.
        /// </summary>
        public const string DriveRoot = "/drive/root:";

        /// <summary>
        /// Maximum chunk size for file transfers.
        /// </summary>
        public const int MaxChunkSize = 4 * 1024 * 1024;
    }
}
