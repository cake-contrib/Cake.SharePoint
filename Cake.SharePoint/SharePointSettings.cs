using Cake.Core.Tooling;
using System;
using System.Collections.Generic;
using System.Text;

namespace Cake.SharePoint
{
    /// <summary>
    /// Settings class to pass-in all the Sharepoint Online parameters
    /// </summary>
    public class SharePointSettings: ToolSettings
    {
        /// <summary>
        /// This the a Azure AD App ID that you need to configure in the Azure AD Portal
        /// More info can be found <see href="<see href="https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/using-csom-for-dotnet-standard#configuring-an-application-in-azure-ad">HERE</see>">
        /// </summary>
        public string AADAppId { get; set; }
        /// <summary>
        /// The SharePoint Online User Name
        /// </summary>
        public string UserName { get; set; }
        /// <summary>
        /// The SharePoint Online Password
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// The Sharepoint Online URL
        /// </summary>
        public string SharePointURL { get; set; }
        /// <summary>
        /// The name of the library you want to access on your site
        /// </summary>
        public string LibraryName { get; set; }
        /// <summary>
        /// The value in MegaBytes of the chunks uploaded to SharePoint for big files.
        /// The default size is 8MB.
        /// </summary>
        public int fileChunkSizeInMB { get; set; } = 8;
    }
}
