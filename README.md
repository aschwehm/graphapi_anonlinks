# SharePoint Anonymous Links Scanner

A production-ready PowerShell solution for detecting and remediating anonymous links in Microsoft 365 SharePoint sites using proper Microsoft Graph API permissions endpoints.

**Version 3.0** - Enterprise-grade scanner with comprehensive Graph API integration.

## Features

- ✅ **Proper Graph API Usage**: Uses `/permissions` endpoint instead of URL pattern matching
- ✅ **Multiple Authentication Methods**: Interactive, App-only (certificate/secret), Device Code
- ✅ **Full Pagination Support**: Handles large tenants with proper `$skiptoken` and `$top` parameters
- ✅ **Exponential Backoff**: Automatic retry with proper 429/503 throttling handling
- ✅ **Concurrent Processing**: Configurable parallelism with throttle limits
- ✅ **Delta Queries**: Support for incremental scans (optional)
- ✅ **Machine-Readable Output**: JSON and CSV with stable IDs for remediation
- ✅ **Remediation Tooling**: Companion script to delete/modify found links
- ✅ **Comprehensive Testing**: Pester tests with CI/CD ready structure
- ✅ **Enterprise Scalability**: Designed for large tenants with proper batching

## Prerequisites

### Required PowerShell Modules

```powershell
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Sites -Scope CurrentUser  
Install-Module -Name Microsoft.Graph.Files -Scope CurrentUser
Install-Module -Name Pester -Scope CurrentUser  # For testing
```

### Required Permissions

#### For Delegated (Interactive/Device Code) Authentication
- `Sites.Read.All` - Read all SharePoint sites
- `Files.Read.All` - Read all files and permissions
- `User.Read` - Basic user profile

#### For App-Only Authentication
- `Sites.Read.All` - Application permission to read sites
- `Files.Read.All` - Application permission to read files

#### For Remediation
- `Sites.ReadWrite.All` - Modify site permissions
- `Files.ReadWrite.All` - Modify file permissions

### App Registration (for App-Only Authentication)

1. Register an app in Azure AD
2. Add API permissions: `Sites.Read.All`, `Files.Read.All` (Application)
3. Grant admin consent
4. Create certificate or client secret
5. Note the Application (Client) ID and Tenant ID

## Usage

### Basic Interactive Scan

```powershell
.\Find-AnonymousLinks.ps1
```

### App-Only Authentication (Unattended)

```powershell
# Using certificate (recommended)
.\Find-AnonymousLinks.ps1 -AuthMethod AppOnly -ClientId "your-app-id" -TenantId "your-tenant-id" -CertificateThumbprint "cert-thumbprint"

# Using client secret (less secure)
$clientSecret = ConvertTo-SecureString "your-secret" -AsPlainText -Force
.\Find-AnonymousLinks.ps1 -AuthMethod AppOnly -ClientId "your-app-id" -TenantId "your-tenant-id" -ClientSecret $clientSecret
```

### Device Code Authentication (for restricted environments)

```powershell
.\Find-AnonymousLinks.ps1 -AuthMethod DeviceCode
```

### High-Performance Scanning

```powershell
# Scan with high concurrency and JSON output
.\Find-AnonymousLinks.ps1 -MaxConcurrency 10 -BatchSize 50 -OutputFormat JSON

# Filter specific sites
.\Find-AnonymousLinks.ps1 -SiteFilter "Marketing|Sales|Finance"

# Enable delta queries for incremental scans
.\Find-AnonymousLinks.ps1 -EnableDeltaQuery
```

## Detection Method

The scanner uses the **proper Microsoft Graph permissions API** to detect anonymous links:

1. **Enumerates all SharePoint sites** using search with pagination
2. **Gets all drives** for each site
3. **Scans all drive items** with full recursion
4. **Calls `/permissions` endpoint** for each item
5. **Checks `link.scope`** for "anonymous" or "anyone" values
6. **Validates `link.type`** (view, edit, etc.)
7. **Records full metadata** including `permissionId` for remediation

### What is Detected

- **Anonymous sharing links** (`link.scope = "anonymous"`)
- **Anyone with link** permissions (`link.scope = "anyone"`)
- **Link types**: view, edit, embed, blocksDownload, review, etc.
- **Expiration dates** and password protection status
- **Full item hierarchy** with stable IDs

### What is NOT Detected (False Negatives Eliminated)

- ❌ URL pattern matching (fragile and incomplete)
- ❌ Guest users (these are not anonymous)
- ❌ External sharing to specific domains (not anonymous)

## Output

### JSON Output (Recommended)

```json
{
  "scanInfo": {
    "scanDate": "2025-08-25T14:30:00Z",
    "scanDuration": "00:15:23",
    "sitesScanned": 127,
    "drivesScanned": 384,
    "itemsScanned": 15847,
    "anonymousLinksFound": 23,
    "errors": 2,
    "version": "3.0"
  },
  "results": [
    {
      "SiteId": "example.sharepoint.com,site-guid,web-guid",
      "SiteName": "Marketing Team",
      "DriveId": "b!xyz...",
      "ItemId": "item-guid",
      "ItemName": "Q4-Strategy.pptx",
      "ItemPath": "/Documents/Q4-Strategy.pptx",
      "PermissionId": "permission-guid",
      "LinkType": "view",
      "LinkScope": "anonymous",
      "ExpiresOn": "2025-12-31T23:59:59Z",
      "WebUrl": "https://example.sharepoint.com/:p:/s/marketing/abc123",
      "HasPassword": false,
      "Roles": "read"
    }
  ]
}
```

### CSV Output

All same fields as JSON but in tabular format for Excel analysis.

### Remediation Plan

A separate `RemediationPlan_*.json` file is created containing only the data needed for remediation:

```json
{
  "instructions": "Use Remove-AnonymousLinks.ps1 with this file",
  "totalItems": 23,
  "items": [
    {
      "SiteId": "...",
      "DriveId": "...",
      "ItemId": "...",
      "PermissionId": "...",
      "LinkType": "view",
      "LinkScope": "anonymous",
      "ItemName": "Q4-Strategy.pptx",
      "ItemPath": "/Documents/Q4-Strategy.pptx"
    }
  ]
}
```

## Remediation

Use the companion script to fix found issues:

### Preview Changes

```powershell
.\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action Preview
```

### Delete Anonymous Links

```powershell
# Preview first
.\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action Delete -WhatIf

# Execute deletion
.\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action Delete
```

### Convert to Organization-Only Links

```powershell
.\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action ConvertToOrganization
```

### Set Expiration Dates

```powershell
# Set all anonymous links to expire in 30 days
.\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action SetExpiration -ExpirationDays 30
```
## Performance & Scalability

### Large Tenant Support

- **Full pagination**: No arbitrary limits, scans complete tenant
- **Concurrent processing**: Configurable parallelism (default: 5 threads)
- **Exponential backoff**: Proper throttling handling with retry logic
- **Delta queries**: Incremental scans for regular monitoring
- **Batch processing**: Configurable batch sizes for optimal performance

### Performance Characteristics

| Tenant Size | Estimated Runtime | Memory Usage | Recommendations |
|-------------|------------------|--------------|-----------------|
| Small (< 10 sites) | 2-5 minutes | < 100 MB | Default settings |
| Medium (10-100 sites) | 10-30 minutes | 200-500 MB | `-MaxConcurrency 8` |
| Large (100-1000 sites) | 1-3 hours | 500 MB - 2 GB | `-MaxConcurrency 10 -BatchSize 50` |
| Enterprise (> 1000 sites) | 3-8 hours | 1-4 GB | App-only auth, run during off-hours |

### Throttling Protection

- **Automatic retry** with exponential backoff
- **Retry-After header** respect
- **Circuit breaker** for persistent failures
- **Rate limiting** awareness

## Testing

Run the included Pester tests:

```powershell
# Install Pester if needed
Install-Module -Name Pester -Force -SkipPublisherCheck

# Run tests
Invoke-Pester -Path ".\Tests\Find-AnonymousLinks.Tests.ps1" -Verbose
```

Test coverage includes:
- Retry logic and throttling handling
- Pagination functionality
- Permission detection accuracy
- Authentication methods
- Error handling
- Output schema validation

## CI/CD Integration

### PowerShell Script Analyzer

```powershell
# Install PSScriptAnalyzer
Install-Module -Name PSScriptAnalyzer -Force

# Run analysis
Invoke-ScriptAnalyzer -Path ".\Find-AnonymousLinks.ps1" -Severity Error,Warning
```

### Azure DevOps Pipeline Example

```yaml
steps:
- task: PowerShell@2
  displayName: 'Install Dependencies'
  inputs:
    targetType: 'inline'
    script: |
      Install-Module -Name PSScriptAnalyzer -Force
      Install-Module -Name Pester -Force

- task: PowerShell@2
  displayName: 'Run Script Analysis'
  inputs:
    targetType: 'inline'
    script: |
      $results = Invoke-ScriptAnalyzer -Path "Find-AnonymousLinks.ps1" -Severity Error
      if ($results) { throw "Script analysis failed" }

- task: PowerShell@2
  displayName: 'Run Tests'
  inputs:
    targetType: 'inline'
    script: |
      Invoke-Pester -Path "Tests/" -OutputFile "TestResults.xml" -OutputFormat NUnitXml
      
- task: PublishTestResults@2
  inputs:
    testResultsFormat: 'NUnit'
    testResultsFiles: 'TestResults.xml'
```

## Troubleshooting

### Common Issues

**"Authentication failed"**
- Verify app registration permissions
- Check certificate/secret expiration
- Ensure admin consent granted

**"No sites found"**
- Verify `Sites.Read.All` permission
- Check user/app has access to SharePoint
- Try with Global Admin account

**"Throttling errors persist"**
- Reduce `-MaxConcurrency` value
- Increase `-BatchSize` (counter-intuitive but reduces request frequency)
- Run during off-peak hours

**"High memory usage"**
- Reduce `-MaxConcurrency`
- Process in smaller batches with site filtering
- Use delta queries for incremental scans

### Debug Mode

```powershell
# Enable detailed logging
$VerbosePreference = "Continue"
.\Find-AnonymousLinks.ps1 -Verbose
```

## Security Considerations

- **Read-only scanning**: No modifications made during scanning
- **Secure authentication**: Supports certificate-based app authentication
- **Audit logging**: All actions logged with timestamps
- **Principle of least privilege**: Request minimum required permissions
- **Data protection**: Results stored locally, review before sharing

## Limitations

### Current Limitations

- **PowerShell requirement**: Requires PowerShell 5.1 or later
- **Windows bias**: Optimized for Windows, may need adjustments for PowerShell Core on Linux/Mac
- **Graph API dependency**: Requires stable internet connection
- **Permission dependency**: Requires admin-level permissions for comprehensive scanning

### Known Limitations vs. Previous Claims

✅ **Fixed**: No longer uses URL pattern matching  
✅ **Fixed**: No arbitrary 50-item limits  
✅ **Fixed**: Proper throttling and error handling  
✅ **Fixed**: Machine-readable output with stable IDs  
✅ **Fixed**: Remediation tooling included  
✅ **Fixed**: Comprehensive test coverage  

### Future Enhancements

- Support for SharePoint on-premises
- Advanced filtering by site templates
- Integration with Microsoft Purview
- PowerShell Core optimization
- Graph batch API utilization

## Support and Contributing

This is a production-ready security tool. Please:

1. **Test thoroughly** in non-production environments
2. **Review permissions** required for your specific use case
3. **Monitor performance** on large tenants
4. **Report issues** with detailed error information
5. **Contribute improvements** via pull requests

## License

MIT License - See LICENSE file for details.
