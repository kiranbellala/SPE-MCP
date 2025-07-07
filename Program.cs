using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using GraphBetaModels = Microsoft.Graph.Beta.Models;
using GraphBeta = Microsoft.Graph.Beta;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Hosting;
using Microsoft.Graph.Models;

// Custom attribute to inject services into MCP tools
[AttributeUsage(AttributeTargets.Parameter)]
public class FromServicesAttribute : Attribute { }




namespace SpeMcp
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var builder = Host.CreateApplicationBuilder(args);
            builder.Logging.AddConsole(consoleLogOptions =>
            {
                // Configure all logs to go to stderr
                consoleLogOptions.LogToStandardErrorThreshold = Microsoft.Extensions.Logging.LogLevel.Trace;
            });

            // Add configuration for Azure credentials and Graph clients (both regular and Beta)
            builder.Services.AddSingleton<GraphServiceClient>(serviceProvider =>
            {
                var configuration = serviceProvider.GetRequiredService<IConfiguration>();
                var tenantId = configuration["AzureAd:TenantId"];
                var clientId = configuration["AzureAd:ClientId"];
                var clientSecret = configuration["AzureAd:ClientSecret"];

                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                {
                    throw new InvalidOperationException("Azure AD application credentials are missing. Please provide ClientId, ClientSecret, and TenantId in the configuration.");
                }

                // Create credential using client secret
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret, options);

                return new GraphServiceClient(clientSecretCredential, new[] { "https://graph.microsoft.com/.default" });
            });

            // Add Beta Graph Client
            builder.Services.AddSingleton<GraphBeta.GraphServiceClient>(serviceProvider =>
            {
                var configuration = serviceProvider.GetRequiredService<IConfiguration>();
                var tenantId = configuration["AzureAd:TenantId"];
                var clientId = configuration["AzureAd:ClientId"];
                var clientSecret = configuration["AzureAd:ClientSecret"];

                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                {
                    throw new InvalidOperationException("Azure AD application credentials are missing. Please provide ClientId, ClientSecret, and TenantId in the configuration.");
                }

                // Create credential using client secret
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret, options);

                return new GraphBeta.GraphServiceClient(clientSecretCredential, new[] { "https://graph.microsoft.com/.default" });
            });

            // Add HttpClient with authentication for direct Graph API calls
            builder.Services.AddSingleton<HttpClient>(serviceProvider =>
            {
                var configuration = serviceProvider.GetRequiredService<IConfiguration>();
                var tenantId = configuration["AzureAd:TenantId"];
                var clientId = configuration["AzureAd:ClientId"];
                var clientSecret = configuration["AzureAd:ClientSecret"];

                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                {
                    throw new InvalidOperationException("Azure AD application credentials are missing. Please provide ClientId, ClientSecret, and TenantId in the configuration.");
                }

                // Create a client credentials flow to get token
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };

                var clientSecretCredential = new ClientSecretCredential(
                    tenantId, clientId, clientSecret, options);

                // Create authenticated HttpClient that will be used for direct Graph API calls
                var httpClient = new HttpClient();

                // We'll add authentication header in the actual request to ensure token freshness
                return httpClient;
            });

            builder.Services
                .AddMcpServer()
                .WithStdioServerTransport()
                .WithToolsFromAssembly();

            await builder.Build().RunAsync();
        }

        [McpServerToolType]
        public static class SharePointEmbeddedTool
        {
            private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            [McpServerTool, Description("Create a new SharePoint Embedded container (drive) of a specific container type.")]
            public static async Task<string> CreateContainer(
                [FromServices] GraphBeta.GraphServiceClient graphBetaClient,
                [Description("Container Type ID")] string containerTypeId,
                [Description("Display name for the new container")] string displayName)
            {
                try
                {
                    var container = new GraphBetaModels.FileStorageContainer
                    {
                        DisplayName = displayName,
                        ContainerTypeId = Guid.Parse(containerTypeId),
                        Settings = new GraphBetaModels.FileStorageContainerSettings
                        {
                            IsOcrEnabled = false
                        }
                    };

                    var created = await graphBetaClient.Storage.FileStorage.Containers.PostAsync(container);
                    if (created == null)
                    {
                        return "Container creation failed.";
                    }
                    return JsonSerializer.Serialize(created, _jsonOptions);
                }
                catch (Exception ex)
                {
                    return $"Error creating container: {ex.Message}";
                }
                // No extra closing brace here; keep class open for other methods
            }

            [McpServerTool, Description("Add a column (metadata field) to a SharePoint Embedded container (drive).")]
            public static async Task<string> AddColumnToContainer(
                [FromServices] GraphBeta.GraphServiceClient graphBetaClient,
                [Description("Container (Drive) ID")] string containerId,
                [Description("Column display name")] string displayName,
                [Description("Column type (e.g., text, number, boolean)")] string columnType)
            {
                try
                {
                    var column = new GraphBetaModels.ColumnDefinition
                    {
                        DisplayName = displayName,
                    };

                    // Set the column type
                    switch (columnType.ToLowerInvariant())
                    {
                        case "text":
                            column.Text = new GraphBetaModels.TextColumn();
                            break;
                        case "number":
                            column.Number = new GraphBetaModels.NumberColumn();
                            break;
                        case "boolean":
                        case "bool":
                            column.Boolean = new GraphBetaModels.BooleanColumn();
                            break;
                        default:
                            return $"Unsupported column type: {columnType}. Supported types: text, number, boolean.";
                    }

                    // Get the list associated with the drive (container)
                    var list = await graphBetaClient.Drives[containerId].List.GetAsync();
                    if (list == null)
                    {
                        return $"No SharePoint List found for container (drive) ID {containerId}.";
                    }
                    if (list.ParentReference == null || string.IsNullOrEmpty(list.ParentReference.SiteId))
                    {
                        return $"ParentReference or SiteId is missing for the list associated with container (drive) ID {containerId}.";
                    }

                    var createdColumn = await graphBetaClient.Sites[list.ParentReference.SiteId].Lists[list.Id].Columns.PostAsync(column);

                    if (createdColumn == null)
                    {
                        return "Column creation failed.";
                    }

                    return JsonSerializer.Serialize(createdColumn, _jsonOptions);
                }
                catch (Exception ex)
                {
                    return $"Error adding column: {ex.Message}";
                }
            }

            [McpServerTool, Description("Upload a file to a SharePoint Embedded container (drive). Supports large files.")]
            public static async Task<string> UploadFileToContainer(
                [FromServices] GraphServiceClient graphClient,
                [Description("Drive ID or Container ID")] string driveId,
                [Description("Destination folder path within the container (optional, default is root)")] string? folderPath,
                [Description("Local file path")] string localFilePath)
            {
                try
                {
                    if (!System.IO.File.Exists(localFilePath))
                    {
                        return $"File '{localFilePath}' does not exist.";
                    }

                    var fileName = System.IO.Path.GetFileName(localFilePath);
                    byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(localFilePath);
                    using var fileStream = new System.IO.MemoryStream(fileBytes);

                    // Determine the upload path
                    string uploadPath = string.IsNullOrEmpty(folderPath)
                        ? $"/{fileName}"
                        : (folderPath.StartsWith("/") ? folderPath : "/" + folderPath) + $"/{fileName}";

                    // Remove trailing slash if present
                    if (uploadPath.EndsWith("/"))
                    {
                        uploadPath = uploadPath.TrimEnd('/');
                    }

                    // Upload the file using the v5 SDK pattern
                    var uploadedItem = await graphClient.Drives[driveId].Root.ItemWithPath(uploadPath).Content.PutAsync(fileStream);

                    if (uploadedItem != null)
                    {
                        return JsonSerializer.Serialize(uploadedItem, _jsonOptions);
                    }
                    else
                    {
                        return "File upload did not succeed.";
                    }
                }
                catch (Exception ex)
                {
                    return $"Error uploading file: {ex.Message}";
                }
            }

            [McpServerTool, Description("Upload all files from a local folder to a SharePoint Embedded container (drive), preserving folder structure.")]
            public static async Task<string> UploadFolderToContainer(
                [FromServices] GraphServiceClient graphClient,
                [Description("Drive ID or Container ID")] string driveId,
                [Description("Local folder path to upload")] string localFolderPath,
                [Description("Destination folder path within the container (optional, default is root)")] string? destFolderPath = null)
            {
                try
                {
                    if (!System.IO.Directory.Exists(localFolderPath))
                    {
                        return $"Local folder '{localFolderPath}' does not exist.";
                    }
                    var files = System.IO.Directory.GetFiles(localFolderPath, "*", System.IO.SearchOption.AllDirectories);
                    if (files.Length == 0)
                    {
                        return $"No files found in folder '{localFolderPath}'.";
                    }
                    var results = new List<string>();
                    foreach (var filePath in files)
                    {
                        var fileName = System.IO.Path.GetFileName(filePath);
                        var relativePath = System.IO.Path.GetRelativePath(localFolderPath, System.IO.Path.GetDirectoryName(filePath) ?? string.Empty);
                        string uploadPath = string.IsNullOrEmpty(destFolderPath) ? relativePath : System.IO.Path.Combine(destFolderPath, relativePath);
                        uploadPath = uploadPath.Replace("\\", "/").Trim('/');
                        if (!string.IsNullOrEmpty(uploadPath))
                            uploadPath += "/";
                        uploadPath += fileName;
                        byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                        using var fileStream = new System.IO.MemoryStream(fileBytes);
                        var graphUploadPath = $"/{uploadPath}";
                        if (graphUploadPath.EndsWith("/"))
                            graphUploadPath = graphUploadPath.TrimEnd('/');
                        try
                        {
                            var uploadedItem = await graphClient.Drives[driveId].Root.ItemWithPath(graphUploadPath).Content.PutAsync(fileStream);
                            if (uploadedItem != null)
                            {
                                results.Add($"Uploaded: {uploadPath}");
                            }
                            else
                            {
                                results.Add($"Failed: {uploadPath}");
                            }
                        }
                        catch (Exception fileEx)
                        {
                            results.Add($"Error uploading {uploadPath}: {fileEx.Message}");
                        }
                    }
                    return string.Join("\n", results);
                }
                catch (Exception ex)
                {
                    return $"Error uploading folder: {ex.Message}";
                }
            }

            [McpServerTool, Description("Lists SharePoint Embedded containers of a specific container type.")]
            public static async Task<string> ListContainersByType(
                [FromServices] GraphBeta.GraphServiceClient graphBetaClient,
                [Description("Container Type ID")] string containerTypeId)
            {
                try
                {
                    var result = await graphBetaClient.Storage.FileStorage.Containers.GetAsync((requestConfiguration) =>
                    {
                        requestConfiguration.QueryParameters.Filter = $"containerTypeId eq {containerTypeId}";
                    });

                    if (result?.Value == null)
                    {
                        return "No containers found.";
                    }

                    return JsonSerializer.Serialize(result.Value, _jsonOptions);
                }
                catch (Exception ex)
                {
                    return $"Error listing containers: {ex.Message}";
                }
            }

            [McpServerTool, Description("Get details of a specific SharePoint Embedded container.")]
            public static async Task<string> GetContainer(
                [FromServices] GraphBeta.GraphServiceClient graphBetaClient,
                [Description("Container ID")] string containerId)
            {
                try
                {
                    var container = await graphBetaClient.Storage.FileStorage.Containers[containerId].GetAsync();

                    if (container == null)
                    {
                        return $"Container with ID {containerId} not found.";
                    }

                    return JsonSerializer.Serialize(container, _jsonOptions);
                }
                catch (Exception ex)
                {
                    return $"Error getting container details: {ex.Message}";
                }
            }

            [McpServerTool, Description("List files and folders in a SharePoint Embedded container.")]
            public static async Task<string> ListContainerItems(
                [FromServices] GraphServiceClient graphClient,
                [Description("Container ID")] string containerId,
                [Description("Optional folder path within container")] string? folderPath = null)
            {
                try
                {
                    // In SharePoint Embedded, the container ID is the drive ID
                    DriveItemCollectionResponse? result;
                    if (string.IsNullOrEmpty(folderPath))
                    {
                        // Get items from the root folder
                        result = await graphClient.Drives[containerId].Items["root"].Children.GetAsync();
                    }
                    else
                    {
                        try
                        {
                            var pathRequest = folderPath.StartsWith("/") ? folderPath : "/" + folderPath;
                            if (pathRequest.EndsWith("/"))
                            {
                                pathRequest = pathRequest.Substring(0, pathRequest.Length - 1);
                            }
                            var folderItem = await graphClient.Drives[containerId].Root.ItemWithPath(pathRequest).GetAsync();
                            if (folderItem == null)
                            {
                                return $"Folder path '{folderPath}' not found in the container.";
                            }
                            result = await graphClient.Drives[containerId].Items[folderItem.Id].Children.GetAsync();
                        }
                        catch (Exception pathEx)
                        {
                            return $"Error accessing folder path '{folderPath}': {pathEx.Message}";
                        }
                    }
                    if (result?.Value == null || !result.Value.Any())
                    {
                        return "No items found in the specified location.";
                    }
                    return JsonSerializer.Serialize(result.Value, _jsonOptions);
                }
                catch (Exception ex)
                {
                    return $"Error listing container items: {ex.Message}";
                }
            }
        }
    }
}

