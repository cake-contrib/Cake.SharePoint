using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security;
using Cake.Core;
using Cake.Core.Annotations;
using Cake.Core.Diagnostics;
using Microsoft.SharePoint.Client;

namespace Cake.SharePoint
{
    public static class CakeSharepoint
    {
        private static readonly int fileChunkSizeInMB = 10;

        private static Microsoft.SharePoint.Client.Folder GetRemoteFolder(ClientContext ctx, string aRemoteFolder, Microsoft.SharePoint.Client.Folder aRootFolder)
        {
            var folderTree = aRemoteFolder.Split('/');
            var result = aRootFolder;
            foreach (var item in folderTree)
            {
                var tmp = result.Folders.FirstOrDefault(f => f.Name == item);
                if (tmp == null)
                {
                    result.Folders.Add(item);
                    ctx.Load(result.Folders);
                    ctx.ExecuteQueryAsync().Wait();
                    tmp = result.Folders.FirstOrDefault(f => f.Name == item);
                }
                result = tmp;
                ctx.Load(result.Folders);
                ctx.ExecuteQueryAsync().Wait();
            }
            return result;
        }

        [CakeMethodAlias]
        public static void SharePointUploadFile(this ICakeContext cakecontext, string filename, string destinationfoldername, SharePointSettings sharepointdetails)
        {
            // Get the name of the file.
            string uniqueFileName = Path.GetFileName(filename);
            // Get the size of the file.
            long fileSize = new FileInfo(filename).Length;
            cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"Uploading file '{uniqueFileName}' ({(fileSize / 1048576):F} MB) to SharePoint ({destinationfoldername})");
            //Bind to site collection
            var clientcontext = new ClientContext(sharepointdetails.SharePointURL);
            var creds = new SharePointOnlineCredentials(sharepointdetails.UserName, sharepointdetails.Password);
            clientcontext.Credentials = creds;
            //upload file
            var sw = new Stopwatch();
            sw.Start();
            // Each sliced upload requires a unique ID.
            Guid uploadId = Guid.NewGuid();
            // Get the folder to upload into. 
            List docs = clientcontext.Web.Lists.GetByTitle(sharepointdetails.LibraryName);
            clientcontext.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            clientcontext.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            clientcontext.Load(docs.RootFolder.Folders);
            clientcontext.ExecuteQueryAsync().Wait();

            var targetFolder = GetRemoteFolder(clientcontext, destinationfoldername, docs.RootFolder);

            // File object.
            Microsoft.SharePoint.Client.File uploadFile;

            // Calculate block size in bytes.
            int blockSize = fileChunkSizeInMB * 1024 * 1024;

            if (fileSize <= blockSize)
            {
                // Use regular approach.
                using (FileStream fs = new FileStream(filename, FileMode.Open))
                {
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = fs;
                    fileInfo.Url = uniqueFileName;
                    fileInfo.Overwrite = true;
                    uploadFile = targetFolder.Files.Add(fileInfo);
                    clientcontext.Load(uploadFile);
                    clientcontext.ExecuteQueryAsync().Wait();
                    sw.Stop();
                    cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"{DateTime.Now.ToShortTimeString()}: Upload of file '{uniqueFileName}' ({(fileSize / 1048576):F} MB) Finished! ({((fileSize / sw.Elapsed.TotalSeconds) / 1048576):F} MBytes/s)");
                    cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"Upload took {sw.Elapsed} to complete");
                    return;
                }
            }
            else
            {
                // Use large file upload approach.
                ClientResult<long> bytesUploaded = null;

                FileStream fs = null;
                try
                {
                    fs = System.IO.File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(fs))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from file system in blocks. 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // You've reached the end of the file.
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size.
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }
                            cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"Upload @ {Math.Round(((double)totalBytesRead / (double)fileSize) * 100)}% : {(totalBytesRead / 1048576):F}/{(fileSize / 1048576):F} MBytes ({((totalBytesRead / sw.Elapsed.TotalSeconds) / 1048576):F} MBytes/s)");

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = uniqueFileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = targetFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice.
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        clientcontext.ExecuteQueryAsync().Wait();
                                        // fileoffset is the pointer where the next slice will be added.
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // You can only start the upload once.
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to your file.
                                uploadFile = clientcontext.Web.GetFileByServerRelativeUrl(targetFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload.
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        clientcontext.ExecuteQueryAsync().Wait();
                                        return;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload.
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        clientcontext.ExecuteQueryAsync().Wait();
                                        // Update fileoffset for the next slice.
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        }
                    }
                }
                finally
                {
                    sw.Stop();
                    cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"Upload of file '{uniqueFileName}' ({(fileSize / 1048576):F} MB) Finished! ({((fileSize / sw.Elapsed.TotalSeconds) / 1048576):F} MBytes/s)");
                    cakecontext?.Log.Write(Verbosity.Normal, LogLevel.Debug, $"Upload took {sw.Elapsed} to complete");

                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }
            }
        }

        [CakeMethodAlias]
        public static IList<String> SharePointGetFilenamesInFolder(this ICakeContext cakecontext, string destinationfoldername, SharePointSettings sharepointdetails)
        {
            //Bind to site collection
            var clientcontext = new ClientContext(sharepointdetails.SharePointURL);
            var creds = new SharePointOnlineCredentials(sharepointdetails.UserName, sharepointdetails.Password);
            clientcontext.Credentials = creds;
            // Get the folder to upload into. 
            List docs = clientcontext.Web.Lists.GetByTitle(sharepointdetails.LibraryName);
            clientcontext.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            clientcontext.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            clientcontext.Load(docs.RootFolder.Folders);
            clientcontext.ExecuteQueryAsync().Wait();

            var targetFolder = GetRemoteFolder(clientcontext, destinationfoldername, docs.RootFolder);
            clientcontext.Load(targetFolder.Files);
            var result = new List<String>();
            clientcontext.ExecuteQueryAsync().Wait();
            foreach (var filename in targetFolder.Files)
            {
                result.Add(filename.Name);
            }
            return result;
        }

        [CakeMethodAlias]
        public static void SharePointDeleteFilesInFolder(this ICakeContext cakecontext, IList<String> filenames, String destinationfoldername, SharePointSettings sharepointdetails)
        {
            //Bind to site collection
            var clientcontext = new ClientContext(sharepointdetails.SharePointURL);
            var creds = new SharePointOnlineCredentials(sharepointdetails.UserName, sharepointdetails.Password);
            clientcontext.Credentials = creds;
            // Get the folder to upload into. 
            List docs = clientcontext.Web.Lists.GetByTitle(sharepointdetails.LibraryName);
            clientcontext.Load(docs, l => l.RootFolder);
            // Get the information about the folder that will hold the file.
            clientcontext.Load(docs.RootFolder, f => f.ServerRelativeUrl);
            clientcontext.Load(docs.RootFolder.Folders);
            clientcontext.ExecuteQueryAsync().Wait();

            var targetFolder = GetRemoteFolder(clientcontext, destinationfoldername, docs.RootFolder);
            clientcontext.Load(targetFolder.Files);
            var result = new List<String>();
            clientcontext.ExecuteQueryAsync().Wait();
            foreach (var fn in targetFolder.Files)
            {
                if (filenames.Contains(fn.Name))
                {
                    fn.DeleteObject();
                }
            }
            clientcontext.ExecuteQueryAsync().Wait();
        }
    }
}

