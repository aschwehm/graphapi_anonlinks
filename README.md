# SharePoint Anonymous Links Scanner

A PowerShell script that scans Microsoft 365 SharePoint sites for pages and documents containing anonymous access links.

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
2. **Retrieves** all SharePoint sites in your tenant
3. **Scans** each site for:
   - Site pages content
   - Document libraries
   - Embedded links and URLs
4. **Identifies** potential anonymous access links using pattern matching
5. **Reports** findings with detailed information
6. **Exports** results to a CSV file

## Anonymous Link Patterns Detected

The script looks for common patterns that indicate anonymous or guest access:

- `guestaccess=true`
- `anonymous`, `guest`, `anyone`, `public`
- SharePoint sharing link patterns (`:x:/`, `:b:/`, `:f:/`)
- OneDrive short links (`1drv.ms`)
- Microsoft short links (`aka.ms`)
- Generic SharePoint sharing URLs

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
SharePoint Anonymous Links Scanner
=================================
Connecting to Microsoft Graph...
Successfully connected to Microsoft Graph!
Tenant ID: 12345678-1234-1234-1234-123456789012
Account: admin@yourcompany.com

Retrieving SharePoint sites...
Found 25 SharePoint sites
Scanning site: Human Resources
Scanning site: Marketing
WARNING: Found 2 anonymous link(s) in page: Welcome.aspx
Scanning site: Sales

SCAN COMPLETE
=============
Found anonymous links in 1 location(s):

Site: Marketing
Page: Welcome.aspx
URL: https://yourcompany.sharepoint.com/sites/marketing/SitePages/Welcome.aspx
Anonymous Links Found: 2
  - https://yourcompany-my.sharepoint.com/:x:/g/personal/user_yourcompany_com/document123
  - https://1drv.ms/w/s!AbcDefGhiJklMnop

Results exported to: AnonymousLinks_20250811_143022.csv

Script completed!
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
