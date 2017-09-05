/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Sage.Office365.Graph
{
    /// <summary>
    /// Implementation for OneDrive access based on a user principal.
    /// </summary>
    public class OneDrive : SynchronousHandler, IOneDrive
    {
        #region Private constants

        private const string DownloadUrl = "@microsoft.graph.downloadUrl";
        private const string DriveRoot = "/drive/root:";
        private const int MaxChunkSize = 4 * 1024 * 1024;

        #endregion

        #region Private fields

        private IClient _client;
        private User _user;

        #endregion

        #region Private properties

        /// <summary>
        /// Simplification for returning the request for the user's drive.
        /// </summary>
        private IDriveRequestBuilder UserDrive
        {
            get
            {
                _client.SignIn();

                return _client.ServiceClient.Users[_user.Id].Drive;
            }
        }

        #endregion

        #region Internal constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="client">The client interface.</param>
        /// <param name="user">The graph user.</param>
        internal OneDrive(IClient client, User user)
        {
            if (client == null) throw new ArgumentNullException(nameof(client));
            if (user == null) throw new ArgumentNullException(nameof(user));

            _client = client;
            _user = user;
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Formats the drive path by removing any leading "/drive/root:" reference, as well as any
        /// leading '/' character.
        /// </summary>
        /// <param name="pathName">The path name to format.</param>
        /// <returns>The formated path name.</returns>
        private string FormatDrivePath(string pathName)
        {
            if (string.IsNullOrEmpty(pathName)) throw new ArgumentNullException("pathName");

            var temp = pathName.StartsWith(DriveRoot, StringComparison.OrdinalIgnoreCase) ? pathName.Remove(0, DriveRoot.Length) : pathName;

            return (temp.StartsWith("/")) ? temp.Remove(0, 1) : temp;
        }

        /// <summary> 
        /// Reads a file fragment from the stream.
        /// </summary> 
        /// <param name="stream">The file stream to read from.</param> 
        /// <param name="position">The position to start reading from.</param> 
        /// <param name="count">The number of bytes to read.</param> 
        /// <returns>The fragment of file with byte[].</returns> 
        private byte[] ReadFileFragment(Stream stream, long position, int count)
        {
            if ((position >= stream.Length) || (position < 0) || (count <= 0)) return null;

            var trimCount = ((position + count) > stream.Length) ? (stream.Length - position) : count;
            var result = new byte[trimCount];

            stream.Seek(position, SeekOrigin.Begin);
            stream.Read(result, 0, (int)trimCount);

            return result;
        }
        
        #endregion

        #region Public methods

        /// <summary>
        /// Creates a folder under the specified parent folder. If parentFolder is null, then the folder
        /// will be created on the root of the drive.
        /// </summary>
        /// <param name="parentFolder">The parent folder to create the child folder under.</param>
        /// <returns>The drive item for the new folder on success, throws on failure.</returns>
        public DriveItem CreateFolder(DriveItem parentFolder, string folderName)
        {
            if ((parentFolder != null) && (parentFolder.Folder == null)) throw new ArgumentNullException("parentFolder");
            if (string.IsNullOrEmpty(folderName)) throw new ArgumentNullException("folderName");

            var subFolder = new DriveItem()
            {
                Name = folderName,
                Folder = new Folder()
            };

            return (parentFolder == null) ? ExecuteTask(UserDrive.Root.Children.Request().AddAsync(subFolder)) : ExecuteTask(UserDrive.Items[parentFolder.Id].Children.Request().AddAsync(subFolder));
        }

        /// <summary>
        /// Deletes a file or folder from OneDrive.
        /// </summary>
        /// <param name="folderOrFile">The drive item representing the item to delete.</param>
        public void DeleteItem(DriveItem folderOrFile)
        {
            if (folderOrFile == null) throw new ArgumentNullException("folderOrFile");

            ExecuteTask(UserDrive.Items[folderOrFile.Id].Request().DeleteAsync());
        }

        /// <summary>
        /// Downloads the file from One Drive into the folder specified by localPath. 
        /// </summary>
        /// <param name="file">The One Drive file to download.</param>
        /// <param name="localPath">The local path to download the file to.</param>
        /// <param name="overWriteIfExists">Overwrites an existing file if set to true.</param>
        public void DownloadFile(DriveItem file, string localPath, bool overWriteIfExists = true)
        {
            if ((file == null) || (file.File == null)) throw new ArgumentNullException("file");
            if (string.IsNullOrEmpty(localPath)) throw new ArgumentNullException("localPath");
            if (!Directory.Exists(localPath)) Directory.CreateDirectory(localPath);

            var localFile = Path.Combine(localPath, file.Name);

            if (System.IO.File.Exists(localFile) && !overWriteIfExists) throw new IOException(string.Format("The file '{0}' already exists and overWriteIfExists was set to false.", localFile));

            var removeFile = true;

            try
            {
                using (var stream = new FileStream(localFile, FileMode.Create))
                {
                    using (var content = ExecuteTask(UserDrive.Items[file.Id].Content.Request().GetAsync()))
                    {
                        content.CopyTo(stream);
                        removeFile = false;
                    }
                }
            }
            catch
            {
                if (removeFile && System.IO.File.Exists(localFile)) System.IO.File.Delete(localFile);

                throw;
            }
        }

        /// <summary>
        /// Gets the drive item children for the specified folder drive item. If folder is null,
        /// then the children for the root of the drive are returned.
        /// </summary>
        /// <param name="folder">The folder to obtain the children for.</param>
        /// <returns>The list of drive item children on success, throws on failure.</returns>
        public IList<DriveItem> GetChildren(DriveItem folder = null)
        {
            var childPage = (folder == null) ?  ExecuteTask(UserDrive.Root.Children.Request().GetAsync()) : ExecuteTask(UserDrive.Items[folder.Id].Children.Request().GetAsync());
            var result = new List<DriveItem>();

            while (childPage != null)
            {
                foreach (var item in childPage.CurrentPage) result.Add(item);

                childPage = (childPage.NextPageRequest == null) ? null : ExecuteTask(childPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        /// <summary>
        /// Attempts to get the drive item using the specified file DriveItem.
        /// </summary>
        /// <param name="pathName">The file drive item to obtain.</param>
        /// <returns>The drive item on success, throws on failure.</returns>
        /// <example>
        /// var pdf = Client.GetItem(fileItem);
        /// </example>
        public DriveItem GetItem(DriveItem file)
        {
            if (file == null) throw new ArgumentNullException(nameof(file));

            return GetItem(GetQualifiedPath(file));
        }

        /// <summary>
        /// Attempts to get the drive item using the specified path, which may contain any number of sub folders
        /// seperated by the '/' character.
        /// </summary>
        /// <param name="pathName">The qualified path of the drive item to obtain.</param>
        /// <returns>The drive item on success, throws on failure.</returns>
        /// <example>
        /// var pdf = Client.GetItem("3031414246/SalesHistory/Accounts Receivable Invoice History Report.pdf");
        /// </example>
        public DriveItem GetItem(string pathName)
        {
            var qualifiedName = FormatDrivePath(pathName);

            if (string.IsNullOrEmpty(qualifiedName)) throw new ApplicationException("Formatted path name cannot be null.");

            return ExecuteTask(UserDrive.Root.ItemWithPath(qualifiedName).Request().GetAsync());
        }

        /// <summary>
        /// Returns the parent folder for the specified drive item.
        /// </summary>
        /// <param name="folderOrFile">The item to return the parent for.</param>
        /// <returns>The parent drive item on success, throws on failure.</returns>
        public DriveItem GetParent(DriveItem folderOrFile)
        {
            if (folderOrFile == null) throw new ArgumentNullException("folderOrFile");
            if (folderOrFile.ParentReference == null) return null;

            var temp = FormatDrivePath(folderOrFile.ParentReference.Path);

            return (string.IsNullOrEmpty(temp)) ? ExecuteTask(UserDrive.Root.Request().GetAsync()) : GetItem(temp);
        }

        /// <summary>
        /// Returns the qualified path for the drive item.
        /// </summary>
        /// <param name="folderOrFile">The drive item to get the qualified path for.</param>
        /// <returns>The drive item's qualified path.</returns>
        public string GetQualifiedPath(DriveItem folderOrFile)
        {
            if (folderOrFile == null) throw new ArgumentNullException("item");
            if (folderOrFile.ParentReference == null) return folderOrFile.Name;

            var parent = folderOrFile.ParentReference;
            var temp = FormatDrivePath(parent.Path);

            return (string.IsNullOrEmpty(temp) ? folderOrFile.Name : string.Format("{0}/{1}", temp, folderOrFile.Name));
        }

        /// <summary>
        /// Uploads a local file into the specified One Drive folder. If folder is null, then 
        /// the file will be created on the root of the drive.
        /// </summary>
        /// <param name="folder">The One Drive folder to upload into.</param>
        /// <param name="localFile">The local file to upload.</param>
        /// <returns>The drive item for the newly uploaded file.</returns>
        public DriveItem UploadFile(DriveItem folder, string localFile)
        {
            if ((folder != null) && (folder.Folder == null)) throw new ArgumentNullException("folder");
            if (string.IsNullOrEmpty(localFile)) throw new ArgumentNullException("localFile");
            if (!System.IO.File.Exists(localFile)) throw new FileNotFoundException();

            _client.SignIn();

            using (var stream = new FileStream(localFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (stream.Length < MaxChunkSize)
                {
                    var file = (folder == null) ? ExecuteTask(UserDrive.Root.Children[Path.GetFileName(localFile)].Content.Request().PutAsync<DriveItem>(stream)) : ExecuteTask(UserDrive.Items[folder.Id].Children[Path.GetFileName(localFile)].Content.Request().PutAsync<DriveItem>(stream));

                    return file;
                }

                var position = 0L;
                var totalLength = stream.Length;
                var remotePath = (folder == null) ? null : GetQualifiedPath(folder);
                var remoteFile = (string.IsNullOrEmpty(remotePath) ? Path.GetFileName(localFile) : string.Format("{0}/{1}", remotePath, Path.GetFileName(localFile)));
                var session = ExecuteTask(UserDrive.Root.ItemWithPath(remoteFile).CreateUploadSession().Request().PostAsync());

                while (position < totalLength)
                {
                    var bytes = ReadFileFragment(stream, position, MaxChunkSize);

                    using (var request = new HttpRequestMessage(HttpMethod.Put, session.UploadUrl))
                    {
                        ExecuteTask(_client.Provider.AuthenticateRequestAsync(request));

                        request.Content = new ByteArrayContent(bytes);
                        request.Content.Headers.ContentRange = new ContentRangeHeaderValue(position, (position + bytes.Length) - 1, totalLength);

                        position += bytes.Length;

                        using (var client = new HttpClient())
                        {
                            using (var response = ExecuteTask(client.SendAsync(request)))
                            {
                                response.EnsureSuccessStatusCode();
                            }
                        }
                    }
                }

                return GetItem(remoteFile);
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns the root of the drive.
        /// </summary>
        public DriveItem Root
        {
            get { return ExecuteTask(UserDrive.Root.Request().GetAsync()); }
        }

        #endregion
    }
}
