#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files

<#
.SYNOPSIS
    Simple SharePoint Anonymous Links Scanner - Windows 10 Compatible

.DESCRIPTION
    A simplified version of the SharePoint anonymous links scanner optimized for Windows 10.
    Scans SharePoint sites for anonymous access links with minimal console formatting.

.PARAMETER TenantId
    The tenant ID of your Microsoft 365 organization (optional)

.EXAMPLE
    .\Find-AnonymousLinks-Simple.ps1
    Runs the scanner with basic output

.NOTES
    Author: Generated Script
    Version: 1.0 - Windows 10 Optimized
    Requires: Microsoft.Graph PowerShell SDK
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string[]]$Scope = @(
        "Sites.Read.All",
        "Files.Read.All", 
        "User.Read"
    )
)

# Simple output function for Windows 10 compatibility
function Write-SimpleOutput {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    
    switch ($Level) {
        "Error" { Write-Output "[$timestamp] ERROR: $Message" }
        "Warning" { Write-Output "[$timestamp] WARNING: $Message" }
        "Success" { Write-Output "[$timestamp] SUCCESS: $Message" }
        default { Write-Output "[$timestamp] INFO: $Message" }
    }
}

# Function to test for anonymous links
function Test-AnonymousLink {
    param([string]$Url)
    
    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }
    
    $patterns = @(
        ":x:/", ":b:/", ":f:/", ":p:/", ":w:/",
        "1drv\.ms",
        "guestaccess=true",
        "anonymous",
        "guest",
        "anyone",
        "sharing",
        "-my\.sharepoint\.com/.*:[a-z]:",
        "sharepoint\.com/:[a-z]:"
    )
    
    foreach ($pattern in $patterns) {
        if ($Url -match $pattern) {
            return $true
        }
    }
    
    return $false
}

# Function to scan a site
function Search-SiteSimple {
    param(
        [string]$SiteId,
        [string]$SiteName,
        [string]$SiteUrl
    )
    
    Write-SimpleOutput "Scanning site: $SiteName"
    $results = @()
    
    try {
        # Scan drives and files
        $drives = Get-MgSiteDrive -SiteId $SiteId -All -ErrorAction SilentlyContinue
        
        foreach ($drive in $drives) {
            Write-SimpleOutput "  Checking drive: $($drive.Name)"
            
            try {
                $items = Get-MgDriveItem -DriveId $drive.Id -All -ErrorAction SilentlyContinue
                
                foreach ($item in $items) {
                    if ($item.WebUrl -and (Test-AnonymousLink -Url $item.WebUrl)) {
                        $result = [PSCustomObject]@{
                            SiteName = $SiteName
                            SiteUrl = $SiteUrl
                            ItemName = $item.Name
                            ItemUrl = $item.WebUrl
                            DriveLocation = $drive.Name
                            ScanDate = Get-Date
                        }
                        $results += $result
                        Write-SimpleOutput "    FOUND: Anonymous link in $($item.Name)" -Level "Warning"
                    }
                }
            } catch {
                Write-SimpleOutput "    Error scanning drive: $($_.Exception.Message)" -Level "Error"
            }
        }
    } catch {
        Write-SimpleOutput "  Error scanning site: $($_.Exception.Message)" -Level "Error"
    }
    
    return $results
}

# Main execution
Write-SimpleOutput "SharePoint Anonymous Links Scanner - Simple Version"
Write-SimpleOutput "=================================================="

# Check modules
Write-SimpleOutput "Checking required modules..."
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites", "Microsoft.Graph.Files")
$missingModules = @()

foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    Write-SimpleOutput "Missing modules: $($missingModules -join ', ')" -Level "Error"
    Write-SimpleOutput "Please run Setup.ps1 first" -Level "Error"
    exit 1
}

Write-SimpleOutput "All required modules found" -Level "Success"

try {
    # Connect to Graph
    Write-SimpleOutput "Connecting to Microsoft Graph..."
    
    $connectParams = @{ Scopes = $Scope }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    
    Connect-MgGraph @connectParams -NoWelcome
    
    $context = Get-MgContext
    Write-SimpleOutput "Connected to tenant: $($context.TenantId)" -Level "Success"
    Write-SimpleOutput "Account: $($context.Account)"
    
    # Get sites
    Write-SimpleOutput "Retrieving SharePoint sites..."
    $sites = @()
    
    # Try multiple methods to get sites
    try {
        $sites = Get-MgSite -All -ErrorAction SilentlyContinue
        if ($sites.Count -eq 0) {
            $sites = Get-MgSite -Search "*" -All -ErrorAction SilentlyContinue
        }
    } catch {
        Write-SimpleOutput "Error getting sites: $($_.Exception.Message)" -Level "Error"
    }
    
    Write-SimpleOutput "Found $($sites.Count) sites"
    
    if ($sites.Count -eq 0) {
        Write-SimpleOutput "No sites found. Check permissions." -Level "Warning"
        exit 0
    }
    
    # Scan sites
    $allResults = @()
    $siteCount = 0
    
    foreach ($site in $sites) {
        $siteCount++
        
        $siteName = if ($site.DisplayName) { $site.DisplayName } else { $site.Name }
        $siteUrl = if ($site.WebUrl) { $site.WebUrl } else { $site.Url }
        $siteId = $site.Id
        
        if ($siteName -and $siteUrl -and $siteId) {
            Write-SimpleOutput "[$siteCount/$($sites.Count)] Processing: $siteName"
            
            $siteResults = Search-SiteSimple -SiteId $siteId -SiteName $siteName -SiteUrl $siteUrl
            $allResults += $siteResults
        } else {
            Write-SimpleOutput "[$siteCount/$($sites.Count)] Skipping site with missing data"
        }
    }
    
    # Results
    Write-SimpleOutput ""
    Write-SimpleOutput "SCAN COMPLETE"
    Write-SimpleOutput "============="
    Write-SimpleOutput "Total sites scanned: $siteCount"
    Write-SimpleOutput "Anonymous links found: $($allResults.Count)"
    
    if ($allResults.Count -gt 0) {
        Write-SimpleOutput "" 
        Write-SimpleOutput "ANONYMOUS LINKS FOUND:" -Level "Warning"
        
        foreach ($result in $allResults) {
            Write-SimpleOutput ""
            Write-SimpleOutput "Site: $($result.SiteName)"
            Write-SimpleOutput "Item: $($result.ItemName)"
            Write-SimpleOutput "Location: $($result.DriveLocation)"
            Write-SimpleOutput "URL: $($result.ItemUrl)"
        }
        
        # Export to CSV
        $csvPath = "AnonymousLinks_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $allResults | Export-Csv -Path $csvPath -NoTypeInformation
        Write-SimpleOutput "Results exported to: $csvPath" -Level "Success"
    } else {
        Write-SimpleOutput "No anonymous links found!" -Level "Success"
    }

} catch {
    Write-SimpleOutput "Script error: $($_.Exception.Message)" -Level "Error"
    Write-SimpleOutput "Line: $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-SimpleOutput "Disconnected from Microsoft Graph"
    } catch {
        # Ignore disconnect errors
    }
}

Write-SimpleOutput "Script completed"
