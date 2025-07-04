# Set these values with your actual Azure AD application credentials
# Replace "your-tenant-id", "your-client-id", and "your-client-secret" with your values

dotnet user-secrets set "AzureAd:TenantId" "your-tenant-id" --project c:\_github\SPE-MCP\SpeMcp\SpeMcp.csproj
dotnet user-secrets set "AzureAd:ClientId" "your-client-id" --project c:\_github\SPE-MCP\SpeMcp\SpeMcp.csproj
dotnet user-secrets set "AzureAd:ClientSecret" "your-client-secret" --project c:\_github\SPE-MCP\SpeMcp\SpeMcp.csproj
