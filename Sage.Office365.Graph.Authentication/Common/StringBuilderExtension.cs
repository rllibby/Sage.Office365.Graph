/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using System;
using System.Text;

namespace Sage.Office365.Graph.Authentication.Common
{
    /// <summary>
    /// Extension class for string builder.
    /// </summary>
    public static class StringBuilderExtension
    {
        #region Private constants

        private const char QueryChar = '?';
        private const char AppendChar = '&';
        private const char EqualsChar = '=';

        #endregion

        #region Public methods

        /// <summary>
        /// Appends a name value pair for a uri query.
        /// </summary>
        /// <param name="source">The string builder to append to.</param>
        /// <param name="name">The name for the value pair.</param>
        /// <param name="value">The value for the value pair.</param>
        /// <param name="first">True if the first value pair set.</param>
        public static void AppendQuery(this StringBuilder source, string name, string value, bool first = false)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));

            source.Append(first ? QueryChar : AppendChar);
            source.Append(name);
            source.Append(EqualsChar);
            source.Append(value);
        }

        /// <summary>
        /// Appends a name value pair for a query body.
        /// </summary>
        /// <param name="source">The string builder to append to.</param>
        /// <param name="name">The name for the value pair.</param>
        /// <param name="value">The value for the value pair.</param>
        public static void AppendBody(this StringBuilder source, string name, string value)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));
            if (source.Length > 0) source.Append(AppendChar);

            source.Append(name);
            source.Append(EqualsChar);
            source.Append(value);
        }

        #endregion
    }
}
