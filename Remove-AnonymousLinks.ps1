#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Sites, Microsoft.Graph.Files

<#
.SYNOPSIS
    Removes or modifies anonymous links found by Find-AnonymousLinks.ps1

.DESCRIPTION
    This script processes the output from Find-AnonymousLinks.ps1 and provides options to:
    - Delete anonymous permissions entirely
    - Convert anonymous links to organization-only links
    - Set expiration dates on anonymous links
    - Generate preview of changes without applying them

.PARAMETER RemediationPlanPath
    Path to the remediation plan JSON file created by Find-AnonymousLinks.ps1

.PARAMETER Action
    Action to take: Delete, ConvertToOrganization, SetExpiration, Preview

.PARAMETER ExpirationDays
    Number of days from now to set expiration (for SetExpiration action)

.PARAMETER WhatIf
    Preview changes without applying them

.PARAMETER TenantId
    The tenant ID of your Microsoft 365 organization

.PARAMETER AuthMethod
    Authentication method: Interactive (default), AppOnly, DeviceCode

.PARAMETER ClientId
    Client ID for app-only authentication

.PARAMETER CertificateThumbprint
    Certificate thumbprint for app-only authentication

.PARAMETER BatchSize
    Batch size for remediation operations (default: 10)

.PARAMETER MaxRetries
    Maximum retry attempts for failed requests (default: 3)

.EXAMPLE
    .\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action Preview
    Preview what changes would be made

.EXAMPLE
    .\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action Delete -WhatIf
    Show what permissions would be deleted without actually deleting them

.EXAMPLE
    .\Remove-AnonymousLinks.ps1 -RemediationPlanPath "RemediationPlan_20250825_143530.json" -Action SetExpiration -ExpirationDays 30
    Set all anonymous links to expire in 30 days

.NOTES
    Author: Enhanced Security Scanner
    Version: 1.0
    Requires: Microsoft.Graph PowerShell SDK
    License: MIT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$RemediationPlanPath,
    
    [Parameter(Mandatory = $true)]
    [ValidateSet("Delete", "ConvertToOrganization", "SetExpiration", "Preview")]
    [string]$Action,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$ExpirationDays = 30,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Interactive", "AppOnly", "DeviceCode")]
    [string]$AuthMethod = "Interactive",
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 50)]
    [int]$BatchSize = 10,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3
)

# Configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Required Graph permissions
$DelegatedScopes = @(
    "Sites.ReadWrite.All",
    "Files.ReadWrite.All", 
    "User.Read"
)

# Stats tracking
$script:Stats = @{
    TotalItems = 0
    Processed = 0
    Succeeded = 0
    Failed = 0
    Skipped = 0
    StartTime = Get-Date
}

# Retry configuration
$script:RetryConfig = @{
    MaxRetries = $MaxRetries
    BaseDelaySeconds = 1
    MaxDelaySeconds = 60
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
    Write-Host "$($Level.ToUpper()): $Message" -ForegroundColor $color
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
        
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            $retryAfter = $_.Exception.Response.Headers["Retry-After"]
        }
        
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
        
        throw
    }
}

#endregion

#region Authentication Functions

function Connect-GraphWithMethod {
    param(
        [string]$Method,
        [string]$TenantId,
        [string]$ClientId,
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
                if (-not $ClientId -or -not $TenantId -or -not $CertificateThumbprint) {
                    throw "ClientId, TenantId, and CertificateThumbprint are required for app-only authentication"
                }
                
                $connectParams.ClientId = $ClientId
                $connectParams.TenantId = $TenantId
                $connectParams.CertificateThumbprint = $CertificateThumbprint
                Write-Log "Connecting with app-only authentication..."
            }
        }
        
        Connect-MgGraph @connectParams -NoWelcome
        
        $context = Get-MgContext
        Write-Log "Successfully connected to tenant: $($context.TenantId)" -Level Success
        
        return $true
    }
    catch {
        Write-Log "Authentication failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}

#endregion

#region Remediation Functions

[CmdletBinding(SupportsShouldProcess)]
function Remove-Permission {
    param(
        [PSCustomObject]$Item
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/drives/$($Item.DriveId)/items/$($Item.ItemId)/permissions/$($Item.PermissionId)"
        
        if ($PSCmdlet.ShouldProcess("$($Item.ItemName) - Permission $($Item.PermissionId)", "Delete Permission")) {
            Invoke-GraphWithRetry -Uri $uri -Method "DELETE"
            Write-Log "DELETED: Permission $($Item.PermissionId) from '$($Item.ItemName)'" -Level Success
            return $true
        } else {
            Write-Log "WOULD DELETE: Permission $($Item.PermissionId) from '$($Item.ItemName)'" -Level Warning
            return $true
        }
    }
    catch {
        Write-Log "FAILED to delete permission $($Item.PermissionId) from '$($Item.ItemName)': $($_.Exception.Message)" -Level Error
        return $false
    }
}

[CmdletBinding(SupportsShouldProcess)]
function Convert-ToOrganizationLink {
    param(
        [PSCustomObject]$Item
    )
    
    try {
        # Get current permission details
        $getUri = "https://graph.microsoft.com/v1.0/drives/$($Item.DriveId)/items/$($Item.ItemId)/permissions/$($Item.PermissionId)"
        $currentPermission = Invoke-GraphWithRetry -Uri $getUri
        
        if (-not $currentPermission.link) {
            Write-Log "SKIPPED: '$($Item.ItemName)' - Not a sharing link" -Level Warning
            return $false
        }
        
        # Update permission to organization scope
        $updateUri = "https://graph.microsoft.com/v1.0/drives/$($Item.DriveId)/items/$($Item.ItemId)/permissions/$($Item.PermissionId)"
        $updateBody = @{
            link = @{
                scope = "organization"
                type = $currentPermission.link.type
            }
        }
        
        if ($PSCmdlet.ShouldProcess("$($Item.ItemName) - Permission $($Item.PermissionId)", "Convert to Organization Link")) {
            Invoke-GraphWithRetry -Uri $updateUri -Method "PATCH" -Body $updateBody
            Write-Log "CONVERTED: '$($Item.ItemName)' from anonymous to organization scope" -Level Success
            return $true
        } else {
            Write-Log "WOULD CONVERT: '$($Item.ItemName)' from anonymous to organization scope" -Level Warning
            return $true
        }
    }
    catch {
        Write-Log "FAILED to convert '$($Item.ItemName)': $($_.Exception.Message)" -Level Error
        return $false
    }
}

[CmdletBinding(SupportsShouldProcess)]
function Set-LinkExpiration {
    param(
        [PSCustomObject]$Item,
        [int]$Days
    )
    
    try {
        # Get current permission details
        $getUri = "https://graph.microsoft.com/v1.0/drives/$($Item.DriveId)/items/$($Item.ItemId)/permissions/$($Item.PermissionId)"
        $currentPermission = Invoke-GraphWithRetry -Uri $getUri
        
        if (-not $currentPermission.link) {
            Write-Log "SKIPPED: '$($Item.ItemName)' - Not a sharing link" -Level Warning
            return $false
        }
        
        # Calculate expiration date
        $expirationDate = (Get-Date).AddDays($Days).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        
        # Update permission with expiration
        $updateUri = "https://graph.microsoft.com/v1.0/drives/$($Item.DriveId)/items/$($Item.ItemId)/permissions/$($Item.PermissionId)"
        $updateBody = @{
            link = @{
                expirationDateTime = $expirationDate
            }
        }
        
        if ($PSCmdlet.ShouldProcess("$($Item.ItemName) - Permission $($Item.PermissionId)", "Set Expiration to $Days days")) {
            Invoke-GraphWithRetry -Uri $updateUri -Method "PATCH" -Body $updateBody
            Write-Log "UPDATED: '$($Item.ItemName)' expiration set to $Days days" -Level Success
            return $true
        } else {
            Write-Log "WOULD SET EXPIRATION: '$($Item.ItemName)' to expire in $Days days" -Level Warning
            return $true
        }
    }
    catch {
        Write-Log "FAILED to set expiration for '$($Item.ItemName)': $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Show-Preview {
    param(
        [PSCustomObject[]]$Items
    )
    
    Write-Log "=== REMEDIATION PREVIEW ===" -Level Info
    Write-Log "Total items to process: $($Items.Count)"
    Write-Log "Action: $Action"
    
    if ($Action -eq "SetExpiration") {
        Write-Log "Expiration: $ExpirationDays days from now"
    }
    
    Write-Log ""
    Write-Log "Items by site:"
    $grouped = $Items | Group-Object SiteName | Sort-Object Name
    foreach ($group in $grouped) {
        Write-Log "  $($group.Name): $($group.Count) items"
    }
    
    Write-Log ""
    Write-Log "Items by link scope:"
    $scopeGrouped = $Items | Group-Object LinkScope | Sort-Object Name
    foreach ($group in $scopeGrouped) {
        Write-Log "  $($group.Name): $($group.Count) items"
    }
    
    Write-Log ""
    Write-Log "Sample items:"
    $Items | Select-Object -First 5 | ForEach-Object {
        Write-Log "  - $($_.SiteName): $($_.ItemName) ($($_.LinkScope))"
    }
    
    if ($Items.Count -gt 5) {
        Write-Log "  ... and $($Items.Count - 5) more items"
    }
}

#endregion

#region Main Execution

function Main {
    Write-Log "SharePoint Anonymous Links Remediation Tool v1.0" -Level Success
    Write-Log "================================================="
    Write-Log ""
    
    # Validate remediation plan file
    if (-not (Test-Path $RemediationPlanPath)) {
        Write-Log "Remediation plan file not found: $RemediationPlanPath" -Level Error
        exit 1
    }
    
    Write-Log "Loading remediation plan: $RemediationPlanPath"
    
    try {
        $remediationData = Get-Content $RemediationPlanPath -Raw | ConvertFrom-Json
        $items = $remediationData.items
        
        if (-not $items -or $items.Count -eq 0) {
            Write-Log "No items found in remediation plan" -Level Warning
            exit 0
        }
        
        $script:Stats.TotalItems = $items.Count
        Write-Log "Loaded $($items.Count) items for remediation" -Level Success
    }
    catch {
        Write-Log "Failed to load remediation plan: $($_.Exception.Message)" -Level Error
        exit 1
    }
    
    # Show preview for all actions
    if ($Action -eq "Preview") {
        Show-Preview -Items $items
        return
    }
    
    # Authenticate (only needed for actual remediation)
    $connected = Connect-GraphWithMethod -Method $AuthMethod -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
    
    if (-not $connected) {
        Write-Log "Authentication failed. Exiting." -Level Error
        exit 1
    }
    
    # Show preview before proceeding
    Show-Preview -Items $items
    Write-Log ""
    
    if (-not $WhatIf) {
        $confirmation = Read-Host "Do you want to proceed with remediation? (y/N)"
        if ($confirmation -ne "y" -and $confirmation -ne "Y") {
            Write-Log "Remediation cancelled by user" -Level Warning
            return
        }
    }
    
    Write-Log "Starting remediation process..."
    Write-Log "Action: $Action"
    if ($Action -eq "SetExpiration") {
        Write-Log "Expiration: $ExpirationDays days"
    }
    Write-Log ""
    
    # Process items in batches
    $batches = for ($i = 0; $i -lt $items.Count; $i += $BatchSize) {
        $items[$i..[Math]::Min($i + $BatchSize - 1, $items.Count - 1)]
    }
    
    foreach ($batch in $batches) {
        Write-Log "Processing batch of $($batch.Count) items..."
        
        foreach ($item in $batch) {
            $script:Stats.Processed++
            $success = $false
            
            try {
                switch ($Action) {
                    "Delete" {
                        $success = Remove-Permission -Item $item
                    }
                    "ConvertToOrganization" {
                        $success = Convert-ToOrganizationLink -Item $item
                    }
                    "SetExpiration" {
                        $success = Set-LinkExpiration -Item $item -Days $ExpirationDays
                    }
                }
                
                if ($success) {
                    $script:Stats.Succeeded++
                } else {
                    $script:Stats.Failed++
                }
            }
            catch {
                Write-Log "Error processing '$($item.ItemName)': $($_.Exception.Message)" -Level Error
                $script:Stats.Failed++
            }
            
            # Progress update
            if ($script:Stats.Processed % 10 -eq 0) {
                $percentComplete = ($script:Stats.Processed / $script:Stats.TotalItems) * 100
                Write-Log "Progress: $($script:Stats.Processed)/$($script:Stats.TotalItems) ($($percentComplete.ToString('F1'))%)"
            }
        }
        
        # Brief pause between batches to be nice to the API
        if ($batch -ne $batches[-1]) {
            Start-Sleep -Milliseconds 500
        }
    }
    
    # Generate completion report
    $duration = (Get-Date) - $script:Stats.StartTime
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    Write-Log ""
    Write-Log "=== REMEDIATION COMPLETE ===" -Level Success
    Write-Log "Duration: $($duration.ToString('hh\:mm\:ss'))"
    Write-Log "Total items: $($script:Stats.TotalItems)"
    Write-Log "Processed: $($script:Stats.Processed)"
    Write-Log "Succeeded: $($script:Stats.Succeeded)" -Level Success
    Write-Log "Failed: $($script:Stats.Failed)" -Level $(if ($script:Stats.Failed -gt 0) { "Warning" } else { "Success" })
    
    # Save completion report
    $reportPath = "RemediationReport_$timestamp.json"
    $report = @{
        action = $Action
        completedAt = Get-Date
        duration = $duration.ToString()
        stats = $script:Stats
        parameters = @{
            remediationPlanPath = $RemediationPlanPath
            expirationDays = if ($Action -eq "SetExpiration") { $ExpirationDays } else { $null }
            batchSize = $BatchSize
            whatIf = $WhatIf.IsPresent
        }
    }
    
    $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Log "Remediation report saved: $reportPath" -Level Success
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
finally {
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Microsoft Graph"
    }
    catch {
        # Ignore disconnect errors
    }
}
