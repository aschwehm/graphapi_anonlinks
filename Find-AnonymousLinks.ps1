#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files

<#
.SYNOPSIS
    Scans Microsoft 365 SharePoint sites for pages containing anonymous access links.

.DESCRIPTION
    This script connects to Microsoft Graph API using modern authentication,
    retrieves all SharePoint sites in the tenant, and scans site pages for
    links that provide anonymous access to resources.

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
    Version: 1.0
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

# Function to check if required modules are installed
function Test-RequiredModules {
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Sites", 
        "Microsoft.Graph.Files"
    )
    
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Error "Missing required modules: $($missingModules -join ', ')"
        Write-Host "Please install missing modules using:"
        Write-Host "Install-Module -Name $($missingModules -join ', ') -Scope CurrentUser"
        return $false
    }
    
    return $true
}

# Function to identify potential anonymous links
function Test-AnonymousLink {
    param([string]$Url)
    
    if ([string]::IsNullOrWhiteSpace($Url)) {
        return $false
    }
    
    # Comprehensive patterns for anonymous/guest access links
    $anonymousPatterns = @(
        # Direct sharing parameters
        "guestaccess=true",
        "anonymous",
        "guest",
        "anyone",
        "public",
        "sharing",
        "share",
        
        # SharePoint/OneDrive sharing patterns
        ":x:/",  # Excel/Word documents sharing
        ":b:/",  # Binary/general file sharing
        ":f:/",  # Folder sharing pattern
        ":i:/",  # Image sharing
        ":p:/",  # PowerPoint sharing
        ":w:/",  # Word document sharing
        ":u:/",  # Upload sharing
        ":v:/",  # Video sharing
        
        # SharePoint sharing URLs
        "sharepoint\.com/:[a-z]:",  # Generic SharePoint sharing
        "_layouts/15/guestaccess.aspx",
        "_layouts/15/shareembedded.aspx", 
        "_layouts/15/accessdenied.aspx",
        
        # OneDrive patterns
        "1drv\.ms",  # OneDrive short links
        "onedrive\.live\.com",
        "-my\.sharepoint\.com/.*:[a-z]:",
        
        # Microsoft short links and redirects
        "aka\.ms",   # Microsoft short links
        "bit\.ly",   # Bit.ly short links (often used for sharing)
        "tinyurl\.com",
        "ow\.ly",
        
        # Sharing tokens and IDs
        "authkey=",
        "resid=",
        "ithint=",
        "e=[A-Za-z0-9]{20,}",  # Long encoded sharing tokens
        
        # Query parameters indicating sharing
        "action=embedview",
        "action=edit",
        "action=interactivepreview",
        "sourcedoc=",
        
        # Teams/Groups sharing
        "teams\.microsoft\.com/.*share",
        "groups\.office\.com/.*share",
        
        # Forms and other sharing
        "forms\.office\.com/.*[Rr]",  # Forms with response links
        "forms\.microsoft\.com/.*[Rr]",
        "whiteboard\.microsoft\.com/.*share",
        
        # Wildcard domain patterns for organizational sharing
        "\.sharepoint\.com/.*:[a-z]:",
        "\.sharepoint\.com/sites/.*/[A-Za-z0-9\-]{20,}",
        
        # Additional suspicious patterns
        "embedded=true",
        "nav=eyJ",  # Base64 encoded navigation (often in sharing)
        "cid=[a-fA-F0-9]{8,}",  # Correlation IDs in sharing URLs
    )
    
    foreach ($pattern in $anonymousPatterns) {
        if ($Url -match $pattern) {
            Write-Verbose "        Anonymous pattern matched: $pattern in URL: $Url"
            return $true
        }
    }
    
    return $false
}

# Function to extract links from page content
function Get-LinksFromContent {
    param([string]$Content)
    
    if ([string]::IsNullOrWhiteSpace($Content)) {
        return @()
    }
    
    $links = @()
    
    # Extract href attributes from HTML content
    $hrefPattern = 'href\s*=\s*["\x27"]([^"\x27]+)["\x27"]'
    $matches = [regex]::Matches($Content, $hrefPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    foreach ($match in $matches) {
        $url = $match.Groups[1].Value
        if (Test-AnonymousLink -Url $url) {
            $links += $url
        }
    }
    
    # Also check for plain URLs in content
    $urlPattern = 'https?://[^\s<>"''()[\]]+'
    $urlMatches = [regex]::Matches($Content, $urlPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    foreach ($match in $urlMatches) {
        $url = $match.Value
        if (Test-AnonymousLink -Url $url) {
            $links += $url
        }
    }
    
    return $links | Select-Object -Unique
}

# Function to check file/folder sharing permissions
function Test-ItemSharingPermissions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveId,
        [Parameter(Mandatory = $true)]
        [string]$ItemId,
        [Parameter(Mandatory = $true)]
        [string]$ItemName,
        [Parameter(Mandatory = $true)]
        [string]$ItemUrl
    )
    
    $sharingInfo = @()
    
    try {
        Write-Verbose "      Checking sharing permissions for: $ItemName"
        
        # Get sharing permissions for the item
        $permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -All -ErrorAction SilentlyContinue
        
        if ($permissions) {
            Write-Verbose "      Found $($permissions.Count) permission(s) for $ItemName"
            
            foreach ($permission in $permissions) {
                # Check for anonymous/guest access indicators
                $isAnonymous = $false
                $reason = ""
                
                # Check permission details
                if ($permission.Link) {
                    Write-Verbose "        Permission has sharing link"
                    
                    # Anonymous link indicators
                    if ($permission.Link.Type -eq "anonymous" -or 
                        $permission.Link.Scope -eq "anonymous" -or
                        $permission.Link.PreventDownload -eq $false) {
                        $isAnonymous = $true
                        $reason = "Anonymous sharing link"
                    }
                    
                    # Check if link allows anyone access
                    if ($permission.Link.Scope -eq "organization" -and $permission.Link.Type -eq "view") {
                        $isAnonymous = $true
                        $reason = "Organization-wide view link"
                    }
                    
                    # Check the actual sharing URL
                    if ($permission.Link.WebUrl -and (Test-AnonymousLink -Url $permission.Link.WebUrl)) {
                        $isAnonymous = $true
                        $reason = "Sharing URL contains anonymous patterns"
                    }
                }
                
                # Check for guest user permissions
                if ($permission.GrantedTo -and $permission.GrantedTo.User) {
                    $userEmail = $permission.GrantedTo.User.Email
                    if ($userEmail -and $userEmail -like "*#EXT#*") {
                        $isAnonymous = $true
                        $reason = "Guest user access: $userEmail"
                    }
                }
                
                # Check for "Anyone" permissions
                if ($permission.Roles -contains "read" -and !$permission.GrantedTo) {
                    $isAnonymous = $true
                    $reason = "Anyone with link can access"
                }
                
                if ($isAnonymous) {
                    $sharingInfo += [PSCustomObject]@{
                        ItemName = $ItemName
                        ItemUrl = $ItemUrl
                        PermissionId = $permission.Id
                        Reason = $reason
                        Roles = $permission.Roles -join ", "
                        LinkType = $permission.Link.Type
                        LinkScope = $permission.Link.Scope
                        ShareLink = $permission.Link.WebUrl
                        GrantedTo = if ($permission.GrantedTo.User.Email) { $permission.GrantedTo.User.Email } else { "Anyone" }
                    }
                    
                    Write-Verbose "        FOUND ANONYMOUS ACCESS: $reason"
                }
            }
        } else {
            Write-Verbose "      No permissions found for $ItemName"
        }
        
    } catch {
        Write-Verbose "      Error checking permissions for $ItemName`: $($_.Exception.Message)"
    }
    
    return $sharingInfo
}
}

# Function to scan a SharePoint site for anonymous links
function Search-SiteForAnonymousLinks {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,
        [Parameter(Mandatory = $true)]
        [string]$SiteName,
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )
    
    Write-Verbose "    Starting detailed scan of site: $SiteName"
    Write-Verbose "    Site URL: $SiteUrl"
    Write-Verbose "    Site ID: $SiteId"
    $results = @()
    
    try {
        # Get site pages
        Write-Verbose "    Getting site pages..."
        $pages = Get-MgSitePage -SiteId $SiteId -All -ErrorAction SilentlyContinue
        
        if ($pages) {
            Write-Verbose "    Found $($pages.Count) page(s) to scan"
            $pageCount = 0
            
            foreach ($page in $pages) {
                $pageCount++
                Write-Verbose "    [$pageCount/$($pages.Count)] Checking page: $($page.Name)"
                
                try {
                    # Get page content if available
                    $pageContent = ""
                    $webPartCount = 0
                    
                    # Try to get the page content
                    if ($page.WebParts) {
                        Write-Verbose "      Found $($page.WebParts.Count) web part(s)"
                        foreach ($webPart in $page.WebParts) {
                            $webPartCount++
                            if ($webPart.InnerHtml) {
                                $pageContent += $webPart.InnerHtml
                                Write-Verbose "        Web part $webPartCount has content ($($webPart.InnerHtml.Length) chars)"
                            } else {
                                Write-Verbose "        Web part $webPartCount has no HTML content"
                            }
                        }
                    } else {
                        Write-Verbose "      No web parts found on this page"
                    }
                    
                    # Check for anonymous links in the content
                    $anonymousLinks = Get-LinksFromContent -Content $pageContent
                    
                    if ($anonymousLinks.Count -gt 0) {
                        $result = [PSCustomObject]@{
                            SiteName = $SiteName
                            SiteUrl = $SiteUrl
                            PageName = $page.Name
                            PageUrl = "$SiteUrl/SitePages/$($page.Name)"
                            AnonymousLinks = $anonymousLinks
                            LinkCount = $anonymousLinks.Count
                            ScanDate = Get-Date
                        }
                        
                        $results += $result
                        Write-Warning "Found $($anonymousLinks.Count) anonymous link(s) in page: $($page.Name)"
                    }
                }
                catch {
                    Write-Verbose "Could not scan page $($page.Name): $($_.Exception.Message)"
                }
            }
        } else {
            Write-Verbose "    No pages found in this site"
        }
        
        # Comprehensive document libraries and files scanning
        Write-Verbose "    Getting document libraries..."
        try {
            $drives = Get-MgSiteDrive -SiteId $SiteId -All -ErrorAction SilentlyContinue
            
            if ($drives) {
                Write-Verbose "    Found $($drives.Count) drive(s) to scan"
                $driveCount = 0
                
                foreach ($drive in $drives) {
                    $driveCount++
                    Write-Verbose "    [$driveCount/$($drives.Count)] Checking drive: $($drive.Name)"
                    
                    try {
                        # Get all items in the drive (files and folders)
                        Write-Verbose "      Getting items from drive..."
                        $items = Get-MgDriveItem -DriveId $drive.Id -All -ErrorAction SilentlyContinue
                        
                        if ($items) {
                            Write-Verbose "      Found $($items.Count) item(s) in drive $($drive.Name)"
                            $itemCount = 0
                            
                            foreach ($item in $items) {
                                $itemCount++
                                Write-Verbose "      [$itemCount/$($items.Count)] Checking item: $($item.Name)"
                                
                                # Check item URL for anonymous patterns
                                if ($item.WebUrl -and (Test-AnonymousLink -Url $item.WebUrl)) {
                                    Write-Verbose "        Found anonymous pattern in item URL"
                                    $result = [PSCustomObject]@{
                                        SiteName = $SiteName
                                        SiteUrl = $SiteUrl
                                        PageName = "File/Folder: $($item.Name) (in $($drive.Name))"
                                        PageUrl = $item.WebUrl
                                        AnonymousLinks = @($item.WebUrl)
                                        LinkCount = 1
                                        ScanDate = Get-Date
                                    }
                                    $results += $result
                                }
                                
                                # Check sharing permissions for the item
                                $sharingInfo = Test-ItemSharingPermissions -DriveId $drive.Id -ItemId $item.Id -ItemName $item.Name -ItemUrl $item.WebUrl
                                
                                if ($sharingInfo.Count -gt 0) {
                                    Write-Verbose "        Found $($sharingInfo.Count) anonymous sharing permission(s)"
                                    
                                    foreach ($sharing in $sharingInfo) {
                                        $result = [PSCustomObject]@{
                                            SiteName = $SiteName
                                            SiteUrl = $SiteUrl
                                            PageName = "Shared Item: $($sharing.ItemName) (in $($drive.Name))"
                                            PageUrl = $sharing.ItemUrl
                                            AnonymousLinks = @($sharing.ShareLink, $sharing.Reason) | Where-Object { $_ }
                                            LinkCount = 1
                                            ScanDate = Get-Date
                                            PermissionDetails = "$($sharing.Reason) - Granted to: $($sharing.GrantedTo)"
                                        }
                                        $results += $result
                                    }
                                }
                                
                                # For performance, limit deep scanning to first 100 items per drive
                                if ($itemCount -ge 100) {
                                    Write-Verbose "      Limiting scan to first 100 items for performance"
                                    break
                                }
                            }
                        } else {
                            Write-Verbose "      No items found in drive $($drive.Name)"
                        }
                        
                    } catch {
                        Write-Verbose "      Could not scan drive $($drive.Name): $($_.Exception.Message)"
                    }
                }
            } else {
                Write-Verbose "    No drives found in this site"
            }
        }
        catch {
            Write-Verbose "    Could not retrieve drives for site $($SiteName): $($_.Exception.Message)"
        }
    }
    catch {
        Write-Error "Error scanning site $($SiteName): $($_.Exception.Message)"
    }
    
    return $results
}

# Main script execution
Write-Host "SharePoint Anonymous Links Scanner" -ForegroundColor Green
Write-Host "=================================" -ForegroundColor Green
Write-Host "Start time: $(Get-Date)" -ForegroundColor Gray
Write-Host ""

# Check for required modules
Write-Host "Checking required modules..." -ForegroundColor Yellow
if (!(Test-RequiredModules)) {
    Write-Host "Required modules are missing. Please run Setup.ps1 first." -ForegroundColor Red
    exit 1
}
Write-Host "All required modules are available." -ForegroundColor Green
Write-Host ""

try {
    # Connect to Microsoft Graph with modern authentication
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Write-Host "Required scopes: $($Scope -join ', ')" -ForegroundColor Gray
    
    $connectParams = @{
        Scopes = $Scope
    }
    
    if ($TenantId) {
        Write-Host "Using specific Tenant ID: $TenantId" -ForegroundColor Gray
        $connectParams.TenantId = $TenantId
    } else {
        Write-Host "Using default tenant (will be determined by authentication)" -ForegroundColor Gray
    }
    
    Write-Host "Opening authentication dialog..." -ForegroundColor Cyan
    Connect-MgGraph @connectParams
    
    Write-Host "Successfully connected to Microsoft Graph!" -ForegroundColor Green
    
    # Get current context
    $context = Get-MgContext
    Write-Host ""
    Write-Host "Connection Details:" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
    Write-Host "Account: $($context.Account)" -ForegroundColor White
    Write-Host "App Name: $($context.AppName)" -ForegroundColor White
    Write-Host "Scopes: $($context.Scopes -join ', ')" -ForegroundColor White
    Write-Host ""
    
    # Get all SharePoint sites
    Write-Host "Retrieving SharePoint sites..." -ForegroundColor Yellow
    Write-Host "This may take a few moments for large tenants..." -ForegroundColor Gray
    
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    # Try multiple methods to get SharePoint sites
    $sites = @()
    
    try {
        Write-Host "Method 1: Attempting to get all sites directly..." -ForegroundColor Gray
        $sites = Get-MgSite -All -ErrorAction SilentlyContinue
        Write-Host "Direct method found $($sites.Count) sites" -ForegroundColor Gray
    } catch {
        Write-Host "Direct method failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # If no sites found, try getting sites by search
    if ($sites.Count -eq 0) {
        try {
            Write-Host "Method 2: Searching for SharePoint sites..." -ForegroundColor Gray
            $searchResults = Get-MgSite -Search "*" -All -ErrorAction SilentlyContinue
            $sites += $searchResults
            Write-Host "Search method found $($sites.Count) additional sites" -ForegroundColor Gray
        } catch {
            Write-Host "Search method failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # If still no sites, try getting the root site and subsites
    if ($sites.Count -eq 0) {
        try {
            Write-Host "Method 3: Getting root site and subsites..." -ForegroundColor Gray
            
            # Get the tenant root site using direct Graph API
            $tenantName = $context.Account.Split('@')[1].Split('.')[0]
            $rootSiteUrl = "https://$tenantName.sharepoint.com"
            
            Write-Host "Attempting to access root site: $rootSiteUrl" -ForegroundColor Gray
            
            # Try to get root site by direct API call
            $rootSiteResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/root" -Method GET -ErrorAction SilentlyContinue
            if ($rootSiteResponse) {
                $sites += $rootSiteResponse
                Write-Host "Found root site: $($rootSiteResponse.displayName)" -ForegroundColor Gray
                
                # Try to get subsites
                try {
                    $subsitesResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($rootSiteResponse.id)/sites" -Method GET -ErrorAction SilentlyContinue
                    if ($subsitesResponse.value) {
                        $sites += $subsitesResponse.value
                        Write-Host "Found $($subsitesResponse.value.Count) subsites" -ForegroundColor Gray
                    }
                } catch {
                    Write-Host "Could not retrieve subsites: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        } catch {
            Write-Host "Root site method failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # If still no sites, try using Graph API directly with different endpoints
    if ($sites.Count -eq 0) {
        try {
            Write-Host "Method 4: Using direct Graph API call..." -ForegroundColor Gray
            $graphResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Method GET
            if ($graphResponse.value) {
                $sites = $graphResponse.value
                Write-Host "Direct Graph API search found $($sites.Count) sites" -ForegroundColor Gray
            }
        } catch {
            Write-Host "Direct Graph API search method failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Try alternative Graph API endpoint
    if ($sites.Count -eq 0) {
        try {
            Write-Host "Method 5: Using alternative Graph API endpoint..." -ForegroundColor Gray
            $graphResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites" -Method GET
            if ($graphResponse.value) {
                $sites = $graphResponse.value
                Write-Host "Alternative Graph API found $($sites.Count) sites" -ForegroundColor Gray
            }
        } catch {
            Write-Host "Alternative Graph API method failed: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    $stopwatch.Stop()
    
    Write-Host "Found $($sites.Count) SharePoint sites (retrieved in $($stopwatch.Elapsed.TotalSeconds.ToString('F2')) seconds)" -ForegroundColor Green
    
    if ($sites.Count -eq 0) {
        Write-Host ""
        Write-Host "NO SHAREPOINT SITES FOUND" -ForegroundColor Red -BackgroundColor DarkRed
        Write-Host "=========================" -ForegroundColor Red -BackgroundColor DarkRed
        Write-Host ""
        Write-Host "Possible reasons:" -ForegroundColor Yellow
        Write-Host "1. Insufficient permissions - You need Sites.Read.All permission" -ForegroundColor White
        Write-Host "2. No SharePoint sites exist in this tenant" -ForegroundColor White
        Write-Host "3. Sites are not accessible with current credentials" -ForegroundColor White
        Write-Host "4. Tenant configuration restricts site discovery" -ForegroundColor White
        Write-Host ""
        Write-Host "Troubleshooting steps:" -ForegroundColor Cyan
        Write-Host "1. Verify you have Global Administrator or SharePoint Administrator rights" -ForegroundColor Gray
        Write-Host "2. Check if SharePoint Online is enabled in your tenant" -ForegroundColor Gray
        Write-Host "3. Try accessing SharePoint sites manually in browser" -ForegroundColor Gray
        Write-Host "4. Contact your Microsoft 365 administrator" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Current permissions granted: $($context.Scopes -join ', ')" -ForegroundColor Gray
        exit 0
    }
    
    # Display site summary
    Write-Host ""
    Write-Host "Site Summary:" -ForegroundColor Cyan
    Write-Host "============" -ForegroundColor Cyan
    $siteTypes = $sites | Group-Object -Property Template | Sort-Object Count -Descending
    foreach ($type in $siteTypes) {
        if ($type.Name) {
            Write-Host "  $($type.Name): $($type.Count) sites" -ForegroundColor Gray
        }
    }
    Write-Host ""
    
    # Initialize results collection
    $allResults = @()
    $totalLinks = 0
    $sitesWithLinks = 0
    
    # Scan each site
    Write-Host "Starting site-by-site scan..." -ForegroundColor Yellow
    Write-Host "=============================" -ForegroundColor Yellow
    Write-Host ""
    
    $siteCount = 0
    $overallStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    foreach ($site in $sites) {
        $siteCount++
        $siteStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Handle different site object types and property names
        $siteName = $site.DisplayName ?? $site.displayName ?? $site.Name ?? $site.name ?? "Unknown Site"
        $siteUrl = $site.WebUrl ?? $site.webUrl ?? $site.Url ?? $site.url ?? ""
        $siteId = $site.Id ?? $site.id ?? ""
        
        # Enhanced progress display
        $percentComplete = ($siteCount / $sites.Count) * 100
        Write-Progress -Activity "Scanning SharePoint Sites for Anonymous Links" -Status "Site $siteCount of $($sites.Count): $siteName" -PercentComplete $percentComplete
        
        Write-Host "[$siteCount/$($sites.Count)] Scanning: " -NoNewline -ForegroundColor Cyan
        Write-Host "$siteName" -ForegroundColor White
        Write-Host "  URL: $siteUrl" -ForegroundColor Gray
        Write-Host "  Site ID: $siteId" -ForegroundColor DarkGray
        
        # Only scan if we have minimum required information
        if (![string]::IsNullOrWhiteSpace($siteUrl) -and ![string]::IsNullOrWhiteSpace($siteName) -and ![string]::IsNullOrWhiteSpace($siteId)) {
            try {
                $siteResults = Search-SiteForAnonymousLinks -SiteId $siteId -SiteName $siteName -SiteUrl $siteUrl
                $allResults += $siteResults
                
                $siteStopwatch.Stop()
                
                if ($siteResults.Count -gt 0) {
                    $sitesWithLinks++
                    $linkCount = ($siteResults | Measure-Object -Property LinkCount -Sum).Sum
                    $totalLinks += $linkCount
                    Write-Host "  Result: " -NoNewline -ForegroundColor Yellow
                    Write-Host "FOUND $linkCount anonymous link(s) in $($siteResults.Count) location(s)" -ForegroundColor Red
                } else {
                    Write-Host "  Result: " -NoNewline -ForegroundColor Yellow
                    Write-Host "No anonymous links found" -ForegroundColor Green
                }
                
                Write-Host "  Scan time: $($siteStopwatch.Elapsed.TotalSeconds.ToString('F2'))s" -ForegroundColor DarkGray
                
            } catch {
                Write-Host "  Result: " -NoNewline -ForegroundColor Yellow
                Write-Host "ERROR - $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "  Scan time: $($siteStopwatch.Elapsed.TotalSeconds.ToString('F2'))s" -ForegroundColor DarkGray
            }
        } else {
            Write-Host "  Result: " -NoNewline -ForegroundColor Yellow
            Write-Host "SKIPPED - Missing required site information" -ForegroundColor Yellow
            Write-Host "    Site Name: $(if($siteName) { 'OK' } else { 'MISSING' })" -ForegroundColor DarkGray
            Write-Host "    Site URL: $(if($siteUrl) { 'OK' } else { 'MISSING' })" -ForegroundColor DarkGray
            Write-Host "    Site ID: $(if($siteId) { 'OK' } else { 'MISSING' })" -ForegroundColor DarkGray
        }
        
        Write-Host ""
        
        # Show intermediate progress every 10 sites
        if ($siteCount % 10 -eq 0) {
            Write-Host "--- Progress Update ---" -ForegroundColor Magenta
            Write-Host "Sites scanned: $siteCount / $($sites.Count)" -ForegroundColor Magenta
            Write-Host "Sites with anonymous links: $sitesWithLinks" -ForegroundColor Magenta
            Write-Host "Total anonymous links found: $totalLinks" -ForegroundColor Magenta
            Write-Host "Elapsed time: $($overallStopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor Magenta
            Write-Host ""
        }
    }
    
    $overallStopwatch.Stop()
    Write-Progress -Activity "Scanning SharePoint Sites" -Completed
    
    # Display comprehensive results
    Write-Host ""
    Write-Host "SCAN COMPLETE" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "=============" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host ""
    
    # Scan Statistics
    Write-Host "Scan Statistics:" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Cyan
    Write-Host "Total sites in tenant: $($sites.Count)" -ForegroundColor White
    Write-Host "Sites successfully scanned: $siteCount" -ForegroundColor White
    Write-Host "Sites with anonymous links: $sitesWithLinks" -ForegroundColor White
    Write-Host "Total anonymous links found: $totalLinks" -ForegroundColor White
    Write-Host "Total scan time: $($overallStopwatch.Elapsed.ToString('hh\:mm\:ss'))" -ForegroundColor White
    Write-Host "Average time per site: $([math]::Round($overallStopwatch.Elapsed.TotalSeconds / $sites.Count, 2))s" -ForegroundColor White
    Write-Host "End time: $(Get-Date)" -ForegroundColor Gray
    Write-Host ""
    
    if ($allResults.Count -eq 0) {
        Write-Host "SUCCESS: No anonymous links found in any SharePoint pages!" -ForegroundColor Green -BackgroundColor DarkGreen
        Write-Host ""
        Write-Host "This indicates good security hygiene in your SharePoint tenant." -ForegroundColor Green
    } else {
        Write-Host "SECURITY ALERT: Found anonymous links in $($allResults.Count) location(s)" -ForegroundColor Red -BackgroundColor DarkRed
        Write-Host ""
        
        # Group results by site for better organization
        $resultsBySite = $allResults | Group-Object -Property SiteName
        
        Write-Host "Detailed Findings:" -ForegroundColor Yellow
        Write-Host "=================" -ForegroundColor Yellow
        
        foreach ($siteGroup in $resultsBySite) {
            Write-Host ""
            Write-Host "SITE: $($siteGroup.Name)" -ForegroundColor Red -BackgroundColor DarkRed
            Write-Host "URL: $($siteGroup.Group[0].SiteUrl)" -ForegroundColor White
            Write-Host "Locations with anonymous links: $($siteGroup.Count)" -ForegroundColor Yellow
            
            $siteTotal = ($siteGroup.Group | Measure-Object -Property LinkCount -Sum).Sum
            Write-Host "Total anonymous links in this site: $siteTotal" -ForegroundColor Yellow
            Write-Host ""
            
            foreach ($result in $siteGroup.Group) {
                Write-Host "  Location: $($result.PageName)" -ForegroundColor Cyan
                Write-Host "  URL: $($result.PageUrl)" -ForegroundColor Gray
                Write-Host "  Anonymous Links Found: $($result.LinkCount)" -ForegroundColor Red
                Write-Host "  Links:" -ForegroundColor White
                
                foreach ($link in $result.AnonymousLinks) {
                    Write-Host "    - $link" -ForegroundColor Yellow
                }
                Write-Host ""
            }
        }
        
        # Export results to CSV
        Write-Host ""
        Write-Host "Exporting Results:" -ForegroundColor Cyan
        Write-Host "=================" -ForegroundColor Cyan
        
        $csvPath = Join-Path $PWD "AnonymousLinks_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $allResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        
        Write-Host "CSV file created: $csvPath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round((Get-Item $csvPath).Length / 1KB, 2)) KB" -ForegroundColor Gray
        Write-Host "Records exported: $($allResults.Count)" -ForegroundColor Gray
        Write-Host ""
        
        # Display CSV column information
        Write-Host "CSV file contains the following columns:" -ForegroundColor Gray
        Write-Host "- SiteName: Name of the SharePoint site" -ForegroundColor Gray
        Write-Host "- SiteUrl: URL of the SharePoint site" -ForegroundColor Gray
        Write-Host "- PageName: Name of the page or location" -ForegroundColor Gray
        Write-Host "- PageUrl: Direct URL to the page/location" -ForegroundColor Gray
        Write-Host "- AnonymousLinks: List of anonymous links found" -ForegroundColor Gray
        Write-Host "- LinkCount: Number of anonymous links in this location" -ForegroundColor Gray
        Write-Host "- ScanDate: Date and time when the scan was performed" -ForegroundColor Gray
        Write-Host ""
        
        # Security recommendations
        Write-Host "Security Recommendations:" -ForegroundColor Red
        Write-Host "========================" -ForegroundColor Red
        Write-Host "1. Review each anonymous link to determine if it's intentional" -ForegroundColor White
        Write-Host "2. Consider restricting sharing permissions where appropriate" -ForegroundColor White
        Write-Host "3. Implement regular monitoring of anonymous sharing" -ForegroundColor White
        Write-Host "4. Educate users about secure sharing practices" -ForegroundColor White
        Write-Host "5. Consider using sensitivity labels to control sharing" -ForegroundColor White
    }
    
} catch {
    Write-Host ""
    Write-Host "ERROR OCCURRED" -ForegroundColor Red -BackgroundColor DarkRed
    Write-Host "==============" -ForegroundColor Red -BackgroundColor DarkRed
    Write-Host ""
    Write-Host "Error details:" -ForegroundColor Red
    Write-Host "Type: $($_.Exception.GetType().Name)" -ForegroundColor White
    Write-Host "Message: $($_.Exception.Message)" -ForegroundColor White
    
    if ($_.Exception.InnerException) {
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor White
    }
    
    Write-Host "Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor White
    Write-Host "Position: $($_.InvocationInfo.PositionMessage)" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Troubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "- Ensure you have the required permissions (Sites.Read.All, Files.Read.All)" -ForegroundColor Gray
    Write-Host "- Check your internet connection" -ForegroundColor Gray
    Write-Host "- Verify that the Microsoft Graph modules are properly installed" -ForegroundColor Gray
    Write-Host "- Try running the script with the -Verbose parameter for more details" -ForegroundColor Gray
    
} finally {
    # Disconnect from Microsoft Graph
    Write-Host ""
    Write-Host "Cleanup:" -ForegroundColor Cyan
    Write-Host "========" -ForegroundColor Cyan
    
    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context) {
            Disconnect-MgGraph
            Write-Host "Successfully disconnected from Microsoft Graph" -ForegroundColor Green
        } else {
            Write-Host "No active Microsoft Graph connection to disconnect" -ForegroundColor Gray
        }
    } catch {
        Write-Host "Note: Could not disconnect cleanly from Microsoft Graph" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "SharePoint Anonymous Links Scanner Done!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Thank you for using the scanner!" -ForegroundColor White
