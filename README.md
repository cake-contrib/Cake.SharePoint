# Cake.SharePoint
An AddOn for Cake to upload files to SharePoint (Office 365) 

## WARNING

starting version 2.0 if the package we use the new Microsoft package to connect to SharePoint Online. This uses a new way to authenticate the user and you will need to make changes to your cake script and make changes on the server.
More info can be found here:
https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/using-csom-for-dotnet-standard#configuring-an-application-in-azure-ad

## Usage

### Include an Add-In directive

#addin "nuget:?package=Cake.SharePoint&loaddependencies=true"

### Upload Task Example

```c#
Task("UploadToSharePoint")
    .IsDependentOn("CreateInstaller")
    .WithCriteria(buildType != "hourly", "Skipped because build is in Hourly Mode")
    .Does(() =>
{    
    var settings = new SharePointSettings
    {
        UserName = sharepointusername,
        Password = sharepointpassword,
        AADAppId = "f4dcf7bb-427a-47fa-9b0e-2f802e7b00ea",
        SharePointURL = "https://yummycheesecompany.sharepoint.com/sites/yummy",
        LibraryName = "Documents"
    };
    
    SharePointUploadFile(installerName, "yummyfiles" , settings);
});
```
