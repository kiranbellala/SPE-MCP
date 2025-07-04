# SharePoint Embedded MCP Server

This is a Model Context Protocol (MCP) server that provides tools to interact with SharePoint via Microsoft Graph API.

## Features

- List SharePoint sites in the tenant
- List document libraries in a SharePoint site
- List contents of a document library or folder

## Prerequisites

- .NET 9.0 SDK
- Azure AD application with appropriate permissions for Microsoft Graph API
- Access to a SharePoint tenant

## Setup

### 1. Register an Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Provide a name for your application
5. Select the appropriate supported account type (single tenant is recommended)
6. Click "Register"
7. Once registered, note the "Application (client) ID" and "Directory (tenant) ID"
8. Create a client secret:
   - Go to "Certificates & secrets"
   - Click "New client secret"
   - Provide a description and select an expiration period
   - Click "Add"
   - Copy the generated client secret value (you won't be able to see it again)

### 2. Configure API Permissions

1. In your app registration, go to "API permissions"
2. Click "Add a permission"
3. Select "Microsoft Graph"
4. Choose "Application permissions"
5. Add the following permissions:
   - Sites.Read.All (to list SharePoint sites)
   - Files.Read.All (to access document libraries and files)
6. Click "Add permissions"
7. Click "Grant admin consent for [your tenant]"

### 3. Configure Authentication Settings

You can store your credentials in user secrets (recommended for development) or in the appsettings.json file.

#### Using User Secrets (Recommended for Development)

Run the following commands from the project directory:

```powershell
dotnet user-secrets set "AzureAd:TenantId" "your-tenant-id"
dotnet user-secrets set "AzureAd:ClientId" "your-client-id"
dotnet user-secrets set "AzureAd:ClientSecret" "your-client-secret"
```

Or use the provided script:

```powershell
# Edit set-secrets.ps1 with your values first
.\set-secrets.ps1
```

#### Using appsettings.json (Not recommended for production)

Update the appsettings.json file with your Azure AD application credentials:

```json
{
  "AzureAd": {
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret"
  }
}
```

## Building and Running

Build the project:

```
dotnet build
```

Run the MCP server:

```
dotnet run
```

## Available MCP Tools

The MCP server provides the following tools:


### SharePoint Embedded Tools

5. `ListContainerTypes`: Lists all available SharePoint Embedded container types.
   - No parameters required

6. `ListContainersByType`: Lists all containers of a specific container type.
   - Parameters: `containerTypeId` (string) - The ID of the container type

7. `GetContainer`: Gets details of a specific SharePoint Embedded container.
   - Parameters: `containerId` (string) - The ID of the container

8. `ListContainerItems`: Lists files and folders in a SharePoint Embedded container.
   - Parameters: 
     - `containerId` (string) - The ID of the container
     - `folderPath` (string, optional) - The path of the folder within the container

## Usage Examples


   ```

### SharePoint Embedded

5. List all container types:
   ```
   ListContainerTypes
   ```

6. List all containers of a specific type (replace with your container type ID):
   ```
   ListContainersByType "12345678-1234-1234-1234-123456789012"
   ```

7. Get details of a specific container (replace with your container ID):
   ```
   GetContainer "12345678-1234-1234-1234-123456789012"
   ```

8. List files and folders in a container (replace with your container ID):
   ```
   ListContainerItems "12345678-1234-1234-1234-123456789012"
   ```

9. List files and folders in a specific folder within a container:
   ```
   ListContainerItems "12345678-1234-1234-1234-123456789012" "Documents/Folder1"
   ```
