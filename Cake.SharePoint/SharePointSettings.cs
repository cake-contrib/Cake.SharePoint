using Cake.Core.Tooling;
using System;
using System.Collections.Generic;
using System.Text;

namespace Cake.SharePoint
{
    public class SharePointSettings: ToolSettings
    {
        public string UserName { get; set; }
        public string Password { get; set; }
        public string SharePointURL { get; set; }
        public string LibraryName { get; set; }
    }
}
