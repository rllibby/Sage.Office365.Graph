/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using System.Collections.Generic;

namespace Sage.Office365.Graph.Interfaces
{
    /// <summary>
    /// Interface definition for OneDrive access.
    /// </summary>
    public interface IOneDrive
    {
        #region Public methods

        /// <summary>
        /// Downloads the file from One Drive into the folder specified by localPath. 
        /// </summary>
        /// <param name="file">The One Drive file to download.</param>
        /// <param name="localPath">The local path to download the file to.</param>
        /// <param name="overWriteIfExists">Overwrites an existing file if set to true.</param>
        void DownloadFile(DriveItem file, string localPath, bool overWriteIfExists = true);

        /// <summary>
        /// Gets the drive item children for the specified folder drive item. If folder is null,
        /// then the children for the root of the drive are returned.
        /// </summary>
        /// <param name="folder">The folder to obtain the children for.</param>
        /// <returns>The list of drive item children on success, throws on failure.</returns>
        IList<DriveItem> GetChildren(DriveItem folder = null);

        /// <summary>
        /// Attempts to get the drive item using the specified path, which may contain any number of sub folders
        /// seperated by the '/' character.
        /// </summary>
        /// <param name="pathName">The qualified path of the drive item to obtain.</param>
        /// <returns>The drive item on success, throws on failure.</returns>
        /// <example>
        /// var pdf = Client.GetItem("3031414246/SalesHistory/Accounts Receivable Invoice History Report.pdf");
        /// </example>
        DriveItem GetItem(string pathName);

        /// <summary>
        /// Returns the parent folder for the specified drive item.
        /// </summary>
        /// <param name="folderOrFile">The item to return the parent for.</param>
        /// <returns>The parent drive item on success, throws on failure.</returns>
        DriveItem GetParent(DriveItem folderOrFile);

        /// <summary>
        /// Returns the qualified path for the drive item.
        /// </summary>
        /// <param name="folderOrFile">The drive item to get the qualified path for.</param>
        /// <returns>The drive item's qualified path.</returns>
        string GetQualifiedPath(DriveItem folderOrFile);

        /// <summary>
        /// Uploads a local file into the specified One Drive folder.
        /// </summary>
        /// <param name="folder">The One Drive folder to upload into.</param>
        /// <param name="localFile">The local file to upload.</param>
        /// <returns>The drive item for the newly uploaded file.</returns>
        DriveItem UploadFile(DriveItem folder, string localFile);

        #endregion
    }
}
