# SharePoint Embedded MCP Server

This is a Model Context Protocol (MCP) server that provides tools to interact with SharePoint Embedded via Microsoft Graph API.



### SharePoint Embedded Tools


 `ListContainersByType`: Lists all containers of a specific ContainerType.
   - Parameters: `containerTypeId` (string) - The ID of the ContainerType

 `GetContainer`: Gets details of a specific SharePoint Embedded container.
   - Parameters: `containerId` (string) - The ID of the container

 `ListContainerItems`: Lists files and folders in a SharePoint Embedded container.
   - Parameters: 
     - `containerId` (string) - The ID of the container
     - `folderPath` (string, optional) - The path of the folder within the container

`UploadFileToContainer`: Uploads a file to a SharePoint Embedded container (drive). Supports large files.
  - Parameters:
    - `driveId` (string) - The ID of the drive or container
    - `folderPath` (string, optional) - The destination folder path within the container (default is root)
    - `localFilePath` (string) - The local file path to upload
