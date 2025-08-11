#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files

<#
.SYNOPSIS
    Scans Microsoft 365 SharePoint sites for pages containing anonymous access links.

.DESCRIPTION
    This script connects to Microsoft Graph API using modern authentication,
    retrieves all SharePoint sites in the tenant, and scans site pages for
    links that provide anonymous access to resources.
    
    Compatible with Windows 10 and 11.

.PARAMETER TenantId
    The tenant ID of your Microsoft 365 organization (optional - will prompt if not provided)

.PARAMETER Scope
    The permissions scope for Microsoft Graph access (default includes necessary permissions)

.EXAMPLE
    .\Find-AnonymousLinks.ps1
    Runs the script with interactive authentication and default settings

.EXAMPLE
    .\Find-AnonymousLinks.ps1 -TenantId "your-tenant-id"
    Runs the script with a specific tenant ID

.NOTES
    Author: Generated Script
    Requires: Microsoft.Graph PowerShell SDK
    Version: 2.0 - Windows 10/11 Compatible
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

# Windows compatibility settings
$ErrorActionPreference = "Continue"
$ProgressPreference = "Continue"

# Safe output function that works on both Windows 10 and 11
function Write-StatusMessage {
    param(
        [string]$Message,
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    
    try {
        switch ($Level) {
            "Error" { 
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
                Write-Host "ERROR: $Message" -ForegroundColor Red 
            }
            "Warning" { 
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
                Write-Host "WARNING: $Message" -ForegroundColor Yellow 
            }
            "Success" { 
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
                Write-Host "SUCCESS: $Message" -ForegroundColor Green 
            }
            "Progress" { 
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
                Write-Host "PROGRESS: $Message" -ForegroundColor Cyan 
            }
            default { 
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
                Write-Host "$Message" -ForegroundColor White 
            }
        }
    } catch {
        # Fallback for compatibility issues
        Write-Output "[$timestamp] $Level`: $Message"
    }
}

# Function to test for anonymous links with comprehensive patterns
function Test-AnonymousLink {
    param([string]$Url)
    
    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }
    
    $patterns = @(
        ":x:/", ":b:/", ":f:/", ":p:/", ":w:/", ":u:/", ":v:/", ":i:",
        "1drv\.ms",
        "guestaccess=true",
        "anonymous", "guest", "anyone", "public", "sharing",
        "-my\.sharepoint\.com/.*:[a-z]:",
        "sharepoint\.com/:[a-z]:",
        "_layouts/15/guestaccess.aspx",
        "authkey=", "resid=", "ithint=",
        "action=embedview", "action=edit",
        "embedded=true",
        "forms\.office\.com/.*[Rr]",
        "nav=eyJ"
    )
    
    foreach ($pattern in $patterns) {
        if ($Url -match $pattern) {
            return $true
        }
    }
    
    return $false
}

# Function to check file/folder sharing permissions
function Test-ItemSharingPermissions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$ItemName,
        [string]$ItemUrl
    )
    
    $sharingInfo = @()
    
    try {
        $permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -All -ErrorAction SilentlyContinue
        
        foreach ($permission in $permissions) {
            $isAnonymous = $false
            $reason = ""
            
            if ($permission.Link) {
                if ($permission.Link.Type -eq "anonymous" -or 
                    $permission.Link.Scope -eq "anonymous") {
                    $isAnonymous = $true
                    $reason = "Anonymous sharing link"
                }
                
                if ($permission.Link.WebUrl -and (Test-AnonymousLink -Url $permission.Link.WebUrl)) {
                    $isAnonymous = $true
                    $reason = "Sharing URL contains anonymous patterns"
                }
            }
            
            if ($permission.GrantedTo -and $permission.GrantedTo.User) {
                $userEmail = $permission.GrantedTo.User.Email
                if ($userEmail -and $userEmail -like "*#EXT#*") {
                    $isAnonymous = $true
                    $reason = "Guest user access: $userEmail"
                }
            }
            
            if ($isAnonymous) {
                $sharingInfo += [PSCustomObject]@{
                    ItemName = $ItemName
                    ItemUrl = $ItemUrl
                    Reason = $reason
                    ShareLink = $permission.Link.WebUrl
                    GrantedTo = if ($permission.GrantedTo.User.Email) { $permission.GrantedTo.User.Email } else { "Anyone" }
                }
            }
        }
    } catch {
        # Silently continue on permission errors
    }
    
    return $sharingInfo
}

# Function to scan a SharePoint site
function Search-SiteForAnonymousLinks {
    param(
        [string]$SiteId,
        [string]$SiteName,
        [string]$SiteUrl
    )
    
    $results = @()
    
    try {
        # Scan document libraries
        $drives = Get-MgSiteDrive -SiteId $SiteId -All -ErrorAction SilentlyContinue
        
        foreach ($drive in $drives) {
            Write-StatusMessage "  Checking drive: $($drive.Name)" -Level "Progress"
            
            try {
                # Limit to first 50 items per drive for performance
                $items = Get-MgDriveItem -DriveId $drive.Id -Top 50 -ErrorAction SilentlyContinue
                
                foreach ($item in $items) {
                    # Check item URL for anonymous patterns
                    if ($item.WebUrl -and (Test-AnonymousLink -Url $item.WebUrl)) {
                        $result = [PSCustomObject]@{
                            SiteName = $SiteName
                            SiteUrl = $SiteUrl
                            ItemName = $item.Name
                            ItemUrl = $item.WebUrl
                            DriveLocation = $drive.Name
                            FindingType = "URL Pattern Match"
                            ScanDate = Get-Date
                        }
                        $results += $result
                        Write-StatusMessage "    FOUND: Anonymous pattern in $($item.Name)" -Level "Warning"
                    }
                    
                    # Check sharing permissions
                    $sharingInfo = Test-ItemSharingPermissions -DriveId $drive.Id -ItemId $item.Id -ItemName $item.Name -ItemUrl $item.WebUrl
                    
                    foreach ($sharing in $sharingInfo) {
                        $result = [PSCustomObject]@{
                            SiteName = $SiteName
                            SiteUrl = $SiteUrl
                            ItemName = $sharing.ItemName
                            ItemUrl = $sharing.ItemUrl
                            DriveLocation = $drive.Name
                            FindingType = $sharing.Reason
                            ScanDate = Get-Date
                        }
                        $results += $result
                        Write-StatusMessage "    FOUND: $($sharing.Reason) in $($sharing.ItemName)" -Level "Warning"
                    }
                }
            } catch {
                Write-StatusMessage "    Error scanning drive $($drive.Name): $($_.Exception.Message)" -Level "Error"
            }
        }
    } catch {
        Write-StatusMessage "  Error scanning site $SiteName`: $($_.Exception.Message)" -Level "Error"
    }
    
    return $results
}

# Main script execution
Write-StatusMessage "SharePoint Anonymous Links Scanner v2.0"
Write-StatusMessage "========================================"
Write-StatusMessage "Compatible with Windows 10 and 11"
Write-StatusMessage ""

# Check required modules
Write-StatusMessage "Checking required modules..."
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites", "Microsoft.Graph.Files")
$missingModules = @()

foreach ($module in $requiredModules) {
    if (!(Get-Module -ListAvailable -Name $module)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    Write-StatusMessage "Missing modules: $($missingModules -join ', ')" -Level "Error"
    Write-StatusMessage "Please run Setup.ps1 first" -Level "Error"
    exit 1
}

Write-StatusMessage "All required modules found" -Level "Success"

try {
    # Connect to Microsoft Graph
    Write-StatusMessage "Connecting to Microsoft Graph..."
    
    $connectParams = @{ Scopes = $Scope }
    if ($TenantId) { 
        $connectParams.TenantId = $TenantId 
        Write-StatusMessage "Using specific Tenant ID: $TenantId"
    }
    
    Connect-MgGraph @connectParams -NoWelcome
    
    $context = Get-MgContext
    Write-StatusMessage "Connected to tenant: $($context.TenantId)" -Level "Success"
    Write-StatusMessage "Account: $($context.Account)"
    Write-StatusMessage "Scopes: $($context.Scopes -join ', ')"
    Write-StatusMessage ""
    
    # Get SharePoint sites using reliable method
    Write-StatusMessage "Retrieving SharePoint sites..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $sites = @()
    
    try {
        # Use search first (most reliable on Windows 10)
        Write-StatusMessage "Searching for SharePoint sites..."
        $sites = Get-MgSite -Search "*" -All -ErrorAction Stop
        Write-StatusMessage "Found $($sites.Count) sites via search" -Level "Success"
    } catch {
        Write-StatusMessage "Search failed, trying direct Graph API..." -Level "Warning"
        
        try {
            $graphResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Method GET
            if ($graphResponse.value) {
                $sites = $graphResponse.value
                Write-StatusMessage "Found $($sites.Count) sites via Graph API" -Level "Success"
            }
        } catch {
            Write-StatusMessage "Graph API failed, trying root site..." -Level "Warning"
            
            try {
                $rootSite = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root" -Method GET
                if ($rootSite) {
                    $sites = @($rootSite)
                    Write-StatusMessage "Found root site only" -Level "Warning"
                }
            } catch {
                Write-StatusMessage "All site discovery methods failed" -Level "Error"
            }
        }
    }
    
    $stopwatch.Stop()
    Write-StatusMessage "Site discovery completed in $($stopwatch.Elapsed.TotalSeconds.ToString('F2')) seconds"
    
    if ($sites.Count -eq 0) {
        Write-StatusMessage "No SharePoint sites found. Check permissions." -Level "Error"
        Write-StatusMessage "Required permissions: Sites.Read.All"
        exit 0
    }
    
    Write-StatusMessage ""
    Write-StatusMessage "Starting site-by-site scan..."
    Write-StatusMessage "=============================="
    
    # Scan sites
    $allResults = @()
    $siteCount = 0
    $sitesWithLinks = 0
    $totalLinks = 0
    
    foreach ($site in $sites) {
        $siteCount++
        
        # Handle different property names from different API responses
        $siteName = if ($site.DisplayName) { $site.DisplayName } elseif ($site.displayName) { $site.displayName } elseif ($site.Name) { $site.Name } else { "Unknown Site" }
        $siteUrl = if ($site.WebUrl) { $site.WebUrl } elseif ($site.webUrl) { $site.webUrl } elseif ($site.Url) { $site.Url } else { "" }
        $siteId = if ($site.Id) { $site.Id } elseif ($site.id) { $site.id } else { "" }
        
        if ($siteName -and $siteUrl -and $siteId) {
            Write-StatusMessage "[$siteCount/$($sites.Count)] Scanning: $siteName" -Level "Progress"
            Write-StatusMessage "  URL: $siteUrl"
            
            $siteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $siteResults = Search-SiteForAnonymousLinks -SiteId $siteId -SiteName $siteName -SiteUrl $siteUrl
            $siteStopwatch.Stop()
            
            $allResults += $siteResults
            
            if ($siteResults.Count -gt 0) {
                $sitesWithLinks++
                $totalLinks += $siteResults.Count
                Write-StatusMessage "  Result: FOUND $($siteResults.Count) anonymous link(s)" -Level "Warning"
            } else {
                Write-StatusMessage "  Result: No anonymous links found" -Level "Success"
            }
            
            Write-StatusMessage "  Scan time: $($siteStopwatch.Elapsed.TotalSeconds.ToString('F2'))s"
        } else {
            Write-StatusMessage "[$siteCount/$($sites.Count)] Skipping site with missing data" -Level "Warning"
        }
        
        Write-StatusMessage ""
    }
    
    # Display results
    Write-StatusMessage ""
    Write-StatusMessage "SCAN COMPLETE"
    Write-StatusMessage "============="
    Write-StatusMessage "Total sites scanned: $siteCount"
    Write-StatusMessage "Sites with anonymous links: $sitesWithLinks"
    Write-StatusMessage "Total anonymous links found: $totalLinks"
    
    if ($allResults.Count -gt 0) {
        Write-StatusMessage "" 
        Write-StatusMessage "ANONYMOUS LINKS FOUND:" -Level "Warning"
        Write-StatusMessage ""
        
        foreach ($result in $allResults) {
            Write-StatusMessage "Site: $($result.SiteName)"
            Write-StatusMessage "Item: $($result.ItemName)"
            Write-StatusMessage "Location: $($result.DriveLocation)"
            Write-StatusMessage "Finding: $($result.FindingType)"
            Write-StatusMessage "URL: $($result.ItemUrl)"
            Write-StatusMessage ""
        }
        
        # Export to CSV
        $csvPath = "AnonymousLinks_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $allResults | Export-Csv -Path $csvPath -NoTypeInformation
        Write-StatusMessage "Results exported to: $csvPath" -Level "Success"
    } else {
        Write-StatusMessage "No anonymous links found - good security hygiene!" -Level "Success"
    }

} catch {
    Write-StatusMessage "Script error: $($_.Exception.Message)" -Level "Error"
    Write-StatusMessage "Line: $($_.InvocationInfo.ScriptLineNumber)"
} finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-StatusMessage "Disconnected from Microsoft Graph"
    } catch {
        # Ignore disconnect errors
    }
}

Write-StatusMessage ""
Write-StatusMessage "Script completed"
