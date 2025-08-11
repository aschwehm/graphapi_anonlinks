# SharePoint Anonymous Links Scanner

A PowerShell script that scans Microsoft 365 SharePoint sites for pages and documents containing anonymous access links.

**Version 2.0** - Universal compatibility with Windows 10 and Windows 11.

## Features

- ✅ **Universal Compatibility**: Works reliably on both Windows 10 and Windows 11
- ✅ **No Hanging Issues**: Optimized site discovery that won't get stuck
- ✅ **Comprehensive Detection**: 50+ patterns for detecting anonymous/guest sharing
- ✅ **Direct Permission Checking**: Scans actual file/folder sharing permissions
- ✅ **Performance Optimized**: Limits scanning to prevent timeouts
- ✅ **Clear Output**: Timestamped progress with safe color handling
- ✅ **Detailed Reporting**: CSV export with categorized findings

## Prerequisites

### Required PowerShell Modules

Install the following Microsoft Graph PowerShell modules:

```powershell
Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module -Name Microsoft.Graph.Sites -Scope CurrentUser  
Install-Module -Name Microsoft.Graph.Files -Scope CurrentUser
```

### Required Permissions

The script requires the following Microsoft Graph permissions:
- `Sites.Read.All` - Read all SharePoint sites
- `Files.Read.All` - Read all files and documents
- `User.Read` - Read basic user profile

## Usage

### Basic Usage

Run the script with interactive authentication:

```powershell
.\Find-AnonymousLinks.ps1
```

### With Specific Tenant ID

```powershell
.\Find-AnonymousLinks.ps1 -TenantId "your-tenant-id-here"
```

### With Verbose Output

```powershell
.\Find-AnonymousLinks.ps1 -Verbose
```

## Authentication

The script uses **modern authentication** with an interactive popup window. When you run the script:

1. A browser window will open for Microsoft 365 authentication
2. Sign in with your admin credentials
3. Grant the requested permissions
4. The script will proceed with the scan

## What the Script Does

1. **Connects** to Microsoft Graph using modern authentication
2. **Discovers** SharePoint sites using reliable methods (optimized for Windows 10/11)
3. **Scans** each site for:
   - Document libraries and files
   - Direct sharing permissions on items
   - Anonymous access patterns in URLs
4. **Identifies** potential anonymous access using:
   - SharePoint sharing link patterns (`:x:/`, `:b:/`, `:f:/`, etc.)
   - OneDrive sharing detection (`1drv.ms`, etc.)
   - Guest user permissions (`#EXT#` users)
   - Anonymous link types and scopes
5. **Reports** findings with detailed categorization
6. **Exports** results to a timestamped CSV file

## Anonymous Link Detection

The script detects various types of anonymous or guest access including:

### URL Patterns
- SharePoint sharing links (`:x:/`, `:b:/`, `:f:/`, `:p:/`, `:w:/`, `:u:/`, `:v:/`, `:i:`)
- OneDrive short links (`1drv.ms`)
- Guest access parameters (`guestaccess=true`, `anonymous`, `guest`, `anyone`)
- Sharing-related URLs (`authkey=`, `resid=`, `embedded=true`)
- Forms and other Office 365 sharing patterns

### Permission Analysis
- Anonymous sharing link permissions
- Guest user access (external users with `#EXT#` in email)
- "Anyone with link" permissions
- Organization-wide sharing links

## Performance & Compatibility

- **Windows 10 Compatible**: Uses reliable site discovery methods that don't hang
- **Windows 11 Optimized**: Takes advantage of newer features when available
- **Performance Limited**: Scans up to 50 items per drive to prevent timeouts
- **Safe Output**: Handles console color compatibility across Windows versions
- **Error Resilient**: Continues scanning even if individual sites fail

## Output

### Console Output

The script provides real-time feedback including:
- Connection status
- Number of sites found
- Progress indicator
- Warning messages for found anonymous links
- Summary of results

### CSV Export

Results are automatically exported to a CSV file named:
`AnonymousLinks_YYYYMMDD_HHMMSS.csv`

The CSV contains:
- Site Name
- Site URL
- Page Name
- Page URL
- Anonymous Links (list)
- Link Count
- Scan Date

## Example Output

```
[14:30:15] SharePoint Anonymous Links Scanner v2.0
[14:30:15] ========================================
[14:30:15] Compatible with Windows 10 and 11
[14:30:16] SUCCESS: All required modules found
[14:30:17] Connecting to Microsoft Graph...
[14:30:20] SUCCESS: Connected to tenant: fa695a3c-5803-4949-bcd1-0eada87cafb4
[14:30:20] Account: admin@yourcompany.com
[14:30:21] SUCCESS: Found 25 sites via search
[14:30:22] Starting site-by-site scan...
[14:30:22] ==============================
[14:30:22] PROGRESS: [1/25] Scanning: Human Resources
[14:30:22] URL: https://yourcompany.sharepoint.com/sites/hr
[14:30:23] SUCCESS: Result: No anonymous links found
[14:30:23] Scan time: 1.23s
[14:30:25] PROGRESS: [2/25] Scanning: Marketing Team
[14:30:25] URL: https://yourcompany.sharepoint.com/sites/marketing
[14:30:25] PROGRESS:   Checking drive: Documents
[14:30:26] WARNING:     FOUND: Anonymous pattern in SharedPresentation.pptx
[14:30:26] WARNING:     FOUND: Guest user access: external@partner.com#EXT# in ProjectPlan.xlsx
[14:30:27] WARNING: Result: FOUND 2 anonymous link(s)
[14:30:27] Scan time: 2.15s

...

[14:35:30] SCAN COMPLETE
[14:35:30] =============
[14:35:30] Total sites scanned: 25
[14:35:30] Sites with anonymous links: 3
[14:35:30] Total anonymous links found: 8

[14:35:30] WARNING: ANONYMOUS LINKS FOUND:

[14:35:30] Site: Marketing Team
[14:35:30] Item: SharedPresentation.pptx
[14:35:30] Location: Documents
[14:35:30] Finding: Anonymous pattern in URL
[14:35:30] URL: https://yourcompany-my.sharepoint.com/:p:/g/personal/user_yourcompany_com/doc123

[14:35:30] Site: Marketing Team
[14:35:30] Item: ProjectPlan.xlsx
[14:35:30] Location: Documents
[14:35:30] Finding: Guest user access: external@partner.com#EXT#
[14:35:30] URL: https://yourcompany.sharepoint.com/sites/marketing/Documents/ProjectPlan.xlsx

[14:35:30] SUCCESS: Results exported to: AnonymousLinks_20250811_143530.csv
[14:35:30] Script completed
```

## Troubleshooting

### Module Installation Issues

If you encounter module installation errors:

```powershell
# Set execution policy (if needed)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Install modules with force
Install-Module -Name Microsoft.Graph.Authentication -Force -Scope CurrentUser
```

### Permission Issues

- Ensure you have **Global Administrator** or **SharePoint Administrator** rights
- Some sites may require additional permissions to access
- The script will continue scanning even if some sites are inaccessible

### Large Tenants

For large tenants with many sites:
- The scan may take considerable time
- Use `-Verbose` to monitor progress
- Consider running during off-hours

## Security Considerations

- The script only **reads** data; it doesn't modify anything
- Authentication tokens are handled securely by the Microsoft Graph SDK
- Results are stored locally in CSV format
- Review the CSV file contents before sharing

## Limitations

- Requires PowerShell 5.1 or later
- May not detect all types of anonymous links (depends on content patterns)
- Some pages with complex formatting may not be fully scanned
- Document content within files is not scanned (only metadata and URLs)

## Support

This script is provided as-is for security auditing purposes. Always test in a non-production environment first.
