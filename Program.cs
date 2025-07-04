//https://laurentkempe.com/2025/03/22/model-context-protocol-made-easy-building-an-mcp-server-in-csharp/

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using GraphBeta = Microsoft.Graph.Beta;
using GraphBetaModels = Microsoft.Graph.Beta.Models;
using ModelContextProtocol.Server;
using System;
using System.ComponentModel;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Azure.Identity;

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
    }


    [McpServerToolType]
    public static class SharePointEmbeddedTool
    {
        private static readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

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
                // Get children items based on the specified path
                DriveItemCollectionResponse? result;
                
                if (string.IsNullOrEmpty(folderPath))
                {
                    // Get items from the root folder
                    result = await graphClient.Drives[containerId].Items["root"].Children.GetAsync();
                }
                else
                {
                    // For folder paths, use the path-based approach
                    try 
                    {
                        // Format the path correctly
                        var pathRequest = folderPath.StartsWith("/") ? folderPath : "/" + folderPath;
                        
                        // Remove trailing slash if present
                        if (pathRequest.EndsWith("/"))
                        {
                            pathRequest = pathRequest.Substring(0, pathRequest.Length - 1);
                        }
                        
                        // Get the folder as an item first
                        var folderItem = await graphClient.Drives[containerId].Root.ItemWithPath(pathRequest).GetAsync();
                        if (folderItem == null)
                        {
                            return $"Folder path '{folderPath}' not found in the container.";
                        }
                        
                        // Get children of the folder using its ID
                        result = await graphClient.Drives[containerId].Items[folderItem.Id].Children.GetAsync();
                    }
                    catch (Exception pathEx)
                    {
                        return $"Error accessing folder path '{folderPath}': {pathEx.Message}";
                    }
                }
                
                // Check if we got any results
                if (result?.Value == null || !result.Value.Any())
                {
                    return "No items found in the specified location.";
                }
                
                // Format the result nicely
                return JsonSerializer.Serialize(result.Value, _jsonOptions);
            }
            catch (Exception ex)
            {
                return $"Error listing container items: {ex.Message}";
            }
        }
    }
}
