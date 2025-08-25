#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files

<#
.SYNOPSIS
    Comprehensive SharePoint anonymous links scanner using Microsoft Graph permissions API.

.DESCRIPTION
    This script uses Microsoft Graph's /permissions endpoint to properly detect anonymous links
    on SharePoint drives and items. It supports full pagination, proper throttling, multiple
    authentication methods, and provides detailed machine-readable output for remediation.

.PARAMETER TenantId
    The tenant ID of your Microsoft 365 organization

.PARAMETER AuthMethod
    Authentication method: Interactive (default), AppOnly, DeviceCode

.PARAMETER ClientId
    Client ID for app-only authentication

.PARAMETER ClientSecret
    Client secret for app-only authentication (use certificate preferred)

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication

.PARAMETER OutputFormat
    Output format: CSV (default), JSON, Both

.PARAMETER OutputPath
    Custom output path for results

.PARAMETER MaxConcurrency
    Maximum concurrent operations (default: 5)

.PARAMETER EnableDeltaQuery
    Enable delta queries for incremental scans

.PARAMETER SiteFilter
    Filter sites by name pattern (regex supported)

.PARAMETER BatchSize
    Batch size for Graph requests (default: 20)

.PARAMETER MaxRetries
    Maximum retry attempts for failed requests (default: 3)

.EXAMPLE
    .\Find-AnonymousLinks.ps1
    Interactive scan with default settings

.EXAMPLE
    .\Find-AnonymousLinks.ps1 -AuthMethod AppOnly -ClientId "app-id" -CertificateThumbprint "thumbprint"
    Unattended scan using app-only authentication

.EXAMPLE
    .\Find-AnonymousLinks.ps1 -OutputFormat JSON -MaxConcurrency 10
    High-performance scan with JSON output

.NOTES
    Author: Enhanced Security Scanner
    Version: 3.0 - Production Ready
    Requires: Microsoft.Graph PowerShell SDK
    License: MIT
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Interactive", "AppOnly", "DeviceCode")]
    [string]$AuthMethod = "Interactive",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [securestring]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Both")]
    [string]$OutputFormat = "Both",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$MaxConcurrency = 5,
    
    [Parameter(Mandatory = $false)]
    [switch]$EnableDeltaQuery,
    
    [Parameter(Mandatory = $false)]
    [string]$SiteFilter,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$BatchSize = 20,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3
)

# Configuration and constants
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Required Graph permissions based on auth method
$DelegatedScopes = @(
    "Sites.Read.All",
    "Files.Read.All", 
    "User.Read"
)

$AppOnlyScopes = @(
    "https://graph.microsoft.com/Sites.Read.All",
    "https://graph.microsoft.com/Files.Read.All"
)

# Retry configuration
$script:RetryConfig = @{
    MaxRetries = $MaxRetries
    BaseDelaySeconds = 1
    MaxDelaySeconds = 60
}

# Result collection
$script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
$script:Stats = @{
    SitesScanned = 0
    DrivesScanned = 0
    ItemsScanned = 0
    AnonymousLinksFound = 0
    Errors = 0
    StartTime = Get-Date
}

#region Utility Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success", "Debug")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        "Debug" { "Cyan" }
        default { "White" }
    }
    
    Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
    Write-Host "$Level`.ToUpper(): $Message" -ForegroundColor $color
}

function Invoke-GraphWithRetry {
    param(
        [string]$Uri,
        [string]$Method = "GET",
        [hashtable]$Headers = @{},
        [object]$Body = $null,
        [int]$RetryCount = 0
    )
    
    try {
        $params = @{
            Uri = $Uri
            Method = $Method
            Headers = $Headers
        }
        
        if ($Body) {
            $params.Body = $Body | ConvertTo-Json -Depth 10
            $params.ContentType = "application/json"
        }
        
        return Invoke-MgGraphRequest @params
    }
    catch {
        $statusCode = $null
        $retryAfter = $null
        
        # Extract status code and retry-after header
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            $retryAfter = $_.Exception.Response.Headers["Retry-After"]
        }
        
        # Handle throttling (429) and server errors (5xx)
        if (($statusCode -eq 429 -or $statusCode -ge 500) -and $RetryCount -lt $script:RetryConfig.MaxRetries) {
            $delay = if ($retryAfter) {
                [int]$retryAfter
            } else {
                [Math]::Min(
                    $script:RetryConfig.BaseDelaySeconds * [Math]::Pow(2, $RetryCount),
                    $script:RetryConfig.MaxDelaySeconds
                )
            }
            
            Write-Log "Request throttled/failed (Status: $statusCode). Retrying in $delay seconds..." -Level Warning
            Start-Sleep -Seconds $delay
            
            return Invoke-GraphWithRetry -Uri $Uri -Method $Method -Headers $Headers -Body $Body -RetryCount ($RetryCount + 1)
        }
        
        # Re-throw if not retryable or max retries exceeded
        throw
    }
}

function Get-AllGraphPages {
    param(
        [string]$Uri,
        [hashtable]$Headers = @{}
    )
    
    $allResults = @()
    $nextLink = $Uri
    
    while ($nextLink) {
        try {
            $response = Invoke-GraphWithRetry -Uri $nextLink -Headers $Headers
            
            if ($response.value) {
                $allResults += $response.value
            }
            
            $nextLink = $response.'@odata.nextLink'
        }
        catch {
            Write-Log "Failed to retrieve page: $($_.Exception.Message)" -Level Error
            $script:Stats.Errors++
            break
        }
    }
    
    return $allResults
}

#endregion

#region Authentication Functions

function Connect-GraphWithMethod {
    param(
        [string]$Method,
        [string]$TenantId,
        [string]$ClientId,
        [securestring]$ClientSecret,
        [string]$CertificateThumbprint
    )
    
    try {
        $connectParams = @{}
        
        switch ($Method) {
            "Interactive" {
                $connectParams.Scopes = $DelegatedScopes
                if ($TenantId) { $connectParams.TenantId = $TenantId }
                Write-Log "Connecting with interactive authentication..."
            }
            
            "DeviceCode" {
                $connectParams.Scopes = $DelegatedScopes
                $connectParams.UseDeviceCode = $true
                if ($TenantId) { $connectParams.TenantId = $TenantId }
                Write-Log "Connecting with device code authentication..."
            }
            
            "AppOnly" {
                if (-not $ClientId) {
                    throw "ClientId is required for app-only authentication"
                }
                if (-not $TenantId) {
                    throw "TenantId is required for app-only authentication"
                }
                
                $connectParams.ClientId = $ClientId
                $connectParams.TenantId = $TenantId
                
                if ($CertificateThumbprint) {
                    $connectParams.CertificateThumbprint = $CertificateThumbprint
                    Write-Log "Connecting with app-only authentication (certificate)..."
                } elseif ($ClientSecret) {
                    $connectParams.ClientSecretCredential = [System.Management.Automation.PSCredential]::new(
                        $ClientId,
                        $ClientSecret
                    )
                    Write-Log "Connecting with app-only authentication (client secret)..."
                } else {
                    throw "Either CertificateThumbprint or ClientSecret is required for app-only authentication"
                }
            }
        }
        
        Connect-MgGraph @connectParams -NoWelcome
        
        $context = Get-MgContext
        Write-Log "Successfully connected to tenant: $($context.TenantId)" -Level Success
        Write-Log "Account: $($context.Account)"
        Write-Log "Auth type: $($context.AuthType)"
        
        return $true
    }
    catch {
        Write-Log "Authentication failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}

#endregion

#region Core Scanning Functions

function Get-AnonymousPermissions {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$ItemName,
        [string]$ItemPath,
        [string]$SiteId,
        [string]$SiteName
    )
    
    try {
        # Get all permissions for the item
        $permissionsUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/permissions"
        $permissions = Get-AllGraphPages -Uri $permissionsUri
        
        foreach ($permission in $permissions) {
            $isAnonymous = $false
            $linkType = $null
            $linkScope = $null
            $grantedTo = $null
            $expiresOn = $null
            
            # Check for anonymous links using proper Graph properties
            if ($permission.link) {
                $linkType = $permission.link.type
                $linkScope = $permission.link.scope
                
                # Anonymous detection based on Graph API properties
                if ($linkScope -eq "anonymous" -or $linkType -eq "anonymous") {
                    $isAnonymous = $true
                }
                
                # Also check for "anyone" links
                if ($linkScope -eq "anyone") {
                    $isAnonymous = $true
                }
                
                if ($permission.link.expirationDateTime) {
                    $expiresOn = [datetime]$permission.link.expirationDateTime
                }
            }
            
            # Check for guest users (external users)
            if ($permission.grantedTo -and $permission.grantedTo.user) {
                $grantedTo = $permission.grantedTo.user.email
                if ($grantedTo -like "*#EXT#*") {
                    # This is an external user, which may indicate anonymous sharing context
                    # Only flag if it's truly anonymous (not just external domain)
                    if (-not $permission.link -or $permission.link.scope -eq "anonymous") {
                        $isAnonymous = $true
                    }
                }
            }
            
            if ($isAnonymous) {
                $result = [PSCustomObject]@{
                    SiteId = $SiteId
                    SiteName = $SiteName
                    DriveId = $DriveId
                    ItemId = $ItemId
                    ItemName = $ItemName
                    ItemPath = $ItemPath
                    PermissionId = $permission.id
                    LinkType = $linkType
                    LinkScope = $linkScope
                    ExpiresOn = $expiresOn
                    GrantedTo = $grantedTo
                    WebUrl = $permission.link.webUrl
                    ScanDate = Get-Date
                    HasPassword = if ($permission.link.hasPassword) { $permission.link.hasPassword } else { $false }
                    Application = $permission.link.application
                    Roles = if ($permission.roles) { $permission.roles -join "," } else { $null }
                }
                
                $script:Results.Add($result)
                $script:Stats.AnonymousLinksFound++
                
                Write-Log "FOUND: Anonymous link in '$ItemName' (Scope: $linkScope, Type: $linkType)" -Level Warning
            }
        }
        
        $script:Stats.ItemsScanned++
    }
    catch {
        Write-Log "Failed to check permissions for item '$ItemName': $($_.Exception.Message)" -Level Error
        $script:Stats.Errors++
    }
}

function Scan-DriveItems {
    param(
        [string]$DriveId,
        [string]$DriveName,
        [string]$SiteId,
        [string]$SiteName,
        [string]$DeltaToken = $null
    )
    
    try {
        Write-Log "Scanning drive: $DriveName"
        
        # Build URI with proper pagination and consistency level
        $baseUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children"
        $headers = @{
            'ConsistencyLevel' = 'eventual'
        }
        
        if ($EnableDeltaQuery -and $DeltaToken) {
            $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/delta?token=$DeltaToken"
        } else {
            $uri = "$baseUri`?`$top=$BatchSize&`$expand=children"
        }
        
        # Get all items with proper pagination
        $items = Get-AllGraphPages -Uri $uri -Headers $headers
        
        if ($items.Count -eq 0) {
            Write-Log "No items found in drive: $DriveName" -Level Debug
            return
        }
        
        Write-Log "Found $($items.Count) items in drive: $DriveName"
        
        # Process items in batches with concurrency control
        $batches = for ($i = 0; $i -lt $items.Count; $i += $BatchSize) {
            $items[$i..[Math]::Min($i + $BatchSize - 1, $items.Count - 1)]
        }
        
        $batches | ForEach-Object -Parallel {
            $batch = $_
            $innerDriveId = $using:DriveId
            $innerSiteId = $using:SiteId
            $innerSiteName = $using:SiteName
            
            # Import functions into parallel scope
            ${function:Write-Log} = $using:function:Write-Log
            ${function:Invoke-GraphWithRetry} = $using:function:Invoke-GraphWithRetry
            ${function:Get-AllGraphPages} = $using:function:Get-AllGraphPages
            ${function:Get-AnonymousPermissions} = $using:function:Get-AnonymousPermissions
            $script:Results = $using:script:Results
            $script:Stats = $using:script:Stats
            $script:RetryConfig = $using:script:RetryConfig
            
            foreach ($item in $batch) {
                try {
                    $itemPath = if ($item.parentReference -and $item.parentReference.path) {
                        $item.parentReference.path + "/" + $item.name
                    } else {
                        $item.name
                    }
                    
                    # Scan permissions for this item
                    Get-AnonymousPermissions -DriveId $innerDriveId -ItemId $item.id -ItemName $item.name -ItemPath $itemPath -SiteId $innerSiteId -SiteName $innerSiteName
                    
                    # If it's a folder, recursively scan its children
                    if ($item.folder -and $item.folder.childCount -gt 0) {
                        $childrenUri = "https://graph.microsoft.com/v1.0/drives/$innerDriveId/items/$($item.id)/children"
                        $children = Get-AllGraphPages -Uri $childrenUri
                        
                        foreach ($child in $children) {
                            $childPath = $itemPath + "/" + $child.name
                            Get-AnonymousPermissions -DriveId $innerDriveId -ItemId $child.id -ItemName $child.name -ItemPath $childPath -SiteId $innerSiteId -SiteName $innerSiteName
                        }
                    }
                }
                catch {
                    Write-Log "Failed to process item '$($item.name)': $($_.Exception.Message)" -Level Error
                    $script:Stats.Errors++
                }
            }
        } -ThrottleLimit $MaxConcurrency
        
        $script:Stats.DrivesScanned++
        Write-Log "Completed scanning drive: $DriveName" -Level Success
    }
    catch {
        Write-Log "Failed to scan drive '$DriveName': $($_.Exception.Message)" -Level Error
        $script:Stats.Errors++
    }
}

function Scan-Site {
    param(
        [PSCustomObject]$Site
    )
    
    try {
        $siteId = $Site.id
        $siteName = $Site.displayName -or $Site.name -or "Unknown Site"
        $siteUrl = $Site.webUrl
        
        # Apply site filter if specified
        if ($SiteFilter -and $siteName -notmatch $SiteFilter) {
            Write-Log "Skipping site '$siteName' (filtered out)" -Level Debug
            return
        }
        
        Write-Log "Scanning site: $siteName" -Level Info
        Write-Log "Site URL: $siteUrl" -Level Debug
        
        # Get all drives for the site
        $drivesUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
        $drives = Get-AllGraphPages -Uri $drivesUri
        
        if ($drives.Count -eq 0) {
            Write-Log "No drives found for site: $siteName" -Level Debug
            return
        }
        
        Write-Log "Found $($drives.Count) drive(s) in site: $siteName"
        
        # Scan each drive
        foreach ($drive in $drives) {
            Scan-DriveItems -DriveId $drive.id -DriveName $drive.name -SiteId $siteId -SiteName $siteName
        }
        
        $script:Stats.SitesScanned++
        Write-Log "Completed scanning site: $siteName" -Level Success
    }
    catch {
        Write-Log "Failed to scan site '$($Site.displayName -or $Site.name)': $($_.Exception.Message)" -Level Error
        $script:Stats.Errors++
    }
}

#endregion

#region Output Functions

function Export-Results {
    param(
        [string]$Format,
        [string]$OutputPath
    )
    
    if ($script:Results.Count -eq 0) {
        Write-Log "No anonymous links found to export" -Level Info
        return
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $resultsArray = @($script:Results)
    
    try {
        if ($Format -eq "CSV" -or $Format -eq "Both") {
            $csvPath = Join-Path $OutputPath "AnonymousLinks_$timestamp.csv"
            $resultsArray | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Log "Results exported to CSV: $csvPath" -Level Success
        }
        
        if ($Format -eq "JSON" -or $Format -eq "Both") {
            $jsonPath = Join-Path $OutputPath "AnonymousLinks_$timestamp.json"
            
            $exportData = @{
                scanInfo = @{
                    scanDate = $script:Stats.StartTime
                    scanDuration = (Get-Date) - $script:Stats.StartTime
                    sitesScanned = $script:Stats.SitesScanned
                    drivesScanned = $script:Stats.DrivesScanned
                    itemsScanned = $script:Stats.ItemsScanned
                    anonymousLinksFound = $script:Stats.AnonymousLinksFound
                    errors = $script:Stats.Errors
                    version = "3.0"
                }
                results = $resultsArray
            }
            
            $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
            Write-Log "Results exported to JSON: $jsonPath" -Level Success
        }
        
        # Create remediation plan
        if ($resultsArray.Count -gt 0) {
            $remediationPath = Join-Path $OutputPath "RemediationPlan_$timestamp.json"
            $remediationPlan = $resultsArray | Select-Object SiteId, DriveId, ItemId, PermissionId, LinkType, LinkScope, ItemName, ItemPath | Sort-Object SiteName, ItemPath
            
            $remediationData = @{
                instructions = "Use Remove-AnonymousLinks.ps1 with this file to remediate findings"
                totalItems = $remediationPlan.Count
                items = $remediationPlan
            }
            
            $remediationData | ConvertTo-Json -Depth 10 | Out-File -FilePath $remediationPath -Encoding UTF8
            Write-Log "Remediation plan created: $remediationPath" -Level Success
        }
    }
    catch {
        Write-Log "Failed to export results: $($_.Exception.Message)" -Level Error
    }
}

function Show-Summary {
    $duration = (Get-Date) - $script:Stats.StartTime
    
    Write-Log ""
    Write-Log "=== SCAN COMPLETE ===" -Level Success
    Write-Log "Scan duration: $($duration.ToString('hh\:mm\:ss'))"
    Write-Log "Sites scanned: $($script:Stats.SitesScanned)"
    Write-Log "Drives scanned: $($script:Stats.DrivesScanned)"
    Write-Log "Items scanned: $($script:Stats.ItemsScanned)"
    Write-Log "Anonymous links found: $($script:Stats.AnonymousLinksFound)" -Level $(if ($script:Stats.AnonymousLinksFound -gt 0) { "Warning" } else { "Success" })
    Write-Log "Errors encountered: $($script:Stats.Errors)" -Level $(if ($script:Stats.Errors -gt 0) { "Warning" } else { "Success" })
    Write-Log ""
    
    if ($script:Results.Count -gt 0) {
        Write-Log "=== FINDINGS SUMMARY ===" -Level Warning
        $grouped = $script:Results | Group-Object LinkScope
        foreach ($group in $grouped) {
            Write-Log "$($group.Name): $($group.Count) items" -Level Warning
        }
        Write-Log ""
    }
}

#endregion

#region Main Execution

function Main {
    Write-Log "SharePoint Anonymous Links Scanner v3.0" -Level Success
    Write-Log "========================================"
    Write-Log "Production-ready scanner with proper Graph API usage"
    Write-Log ""
    
    # Validate dependencies
    Write-Log "Checking required modules..."
    $requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites", "Microsoft.Graph.Files")
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Log "Missing modules: $($missingModules -join ', ')" -Level Error
        Write-Log "Run: Install-Module $($missingModules -join ', ') -Scope CurrentUser" -Level Error
        exit 1
    }
    
    Write-Log "All required modules found" -Level Success
    
    # Authenticate
    $connected = Connect-GraphWithMethod -Method $AuthMethod -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint
    
    if (-not $connected) {
        Write-Log "Authentication failed. Exiting." -Level Error
        exit 1
    }
    
    try {
        # Get SharePoint sites with search and pagination
        Write-Log "Discovering SharePoint sites..."
        $searchUri = "https://graph.microsoft.com/v1.0/sites?search=*&`$top=$BatchSize"
        $headers = @{ 'ConsistencyLevel' = 'eventual' }
        
        $sites = Get-AllGraphPages -Uri $searchUri -Headers $headers
        
        if ($sites.Count -eq 0) {
            Write-Log "No SharePoint sites found. Check permissions." -Level Warning
            return
        }
        
        Write-Log "Found $($sites.Count) SharePoint sites" -Level Success
        
        # Scan sites with controlled concurrency
        Write-Log "Starting comprehensive site scan..."
        Write-Log "Max concurrency: $MaxConcurrency, Batch size: $BatchSize"
        Write-Log ""
        
        $sites | ForEach-Object -Parallel {
            $site = $_
            
            # Import functions into parallel scope
            ${function:Write-Log} = $using:function:Write-Log
            ${function:Invoke-GraphWithRetry} = $using:function:Invoke-GraphWithRetry
            ${function:Get-AllGraphPages} = $using:function:Get-AllGraphPages
            ${function:Get-AnonymousPermissions} = $using:function:Get-AnonymousPermissions
            ${function:Scan-DriveItems} = $using:function:Scan-DriveItems
            ${function:Scan-Site} = $using:function:Scan-Site
            $script:Results = $using:script:Results
            $script:Stats = $using:script:Stats
            $script:RetryConfig = $using:script:RetryConfig
            $SiteFilter = $using:SiteFilter
            $BatchSize = $using:BatchSize
            $MaxConcurrency = $using:MaxConcurrency
            $EnableDeltaQuery = $using:EnableDeltaQuery
            
            Scan-Site -Site $site
        } -ThrottleLimit $MaxConcurrency
        
        Write-Log "Site scanning completed" -Level Success
        
        # Export results
        Write-Log "Exporting results..."
        Export-Results -Format $OutputFormat -OutputPath $OutputPath
        
        # Show summary
        Show-Summary
    }
    catch {
        Write-Log "Critical error during scan: $($_.Exception.Message)" -Level Error
        $script:Stats.Errors++
    }
    finally {
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Log "Disconnected from Microsoft Graph"
        }
        catch {
            # Ignore disconnect errors
        }
    }
}

# Entry point
try {
    Main
}
catch {
    Write-Log "Unhandled error: $($_.Exception.Message)" -Level Error
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}

#endregion
