<#
.SYNOPSIS
    Setup script for SharePoint Anonymous Links Scanner

.DESCRIPTION
    This script installs the required Microsoft Graph PowerShell modules
    and sets up the environment for running the anonymous links scanner.

.PARAMETER Scope
    Installation scope - 'CurrentUser' (default) or 'AllUsers'

.EXAMPLE
    .\Setup.ps1
    Installs modules for current user

.EXAMPLE
    .\Setup.ps1 -Scope AllUsers
    Installs modules for all users (requires admin rights)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$Scope = "CurrentUser"
)

Write-Host "SharePoint Anonymous Links Scanner - Setup" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green

# Required modules
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Sites",
    "Microsoft.Graph.Files"
)

Write-Host "Installation scope: $Scope" -ForegroundColor Cyan

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion
Write-Host "PowerShell version: $($psVersion.ToString())" -ForegroundColor Gray

if ($psVersion.Major -lt 5) {
    Write-Error "PowerShell 5.0 or later is required. Current version: $($psVersion.ToString())"
    exit 1
}

# Set execution policy if needed
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-Host "Current execution policy: $currentPolicy" -ForegroundColor Gray
    
    if ($currentPolicy -eq "Restricted") {
        Write-Host "Setting execution policy to RemoteSigned for CurrentUser..." -ForegroundColor Yellow
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "Execution policy updated!" -ForegroundColor Green
    }
} catch {
    Write-Warning "Could not check/set execution policy: $($_.Exception.Message)"
}

# Install PowerShellGet if needed
try {
    $psGet = Get-Module -ListAvailable PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
    if (!$psGet -or $psGet.Version -lt [Version]"2.0.0") {
        Write-Host "Installing/updating PowerShellGet..." -ForegroundColor Yellow
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope $Scope
        Write-Host "PowerShellGet installed/updated!" -ForegroundColor Green
        Write-Host "Please restart PowerShell and run this setup script again." -ForegroundColor Red
        exit 0
    }
} catch {
    Write-Warning "Could not install PowerShellGet: $($_.Exception.Message)"
}

Write-Host "`nInstalling required modules..." -ForegroundColor Yellow

foreach ($module in $requiredModules) {
    Write-Host "Checking module: $module" -ForegroundColor Cyan
    
    try {
        $installedModule = Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending | Select-Object -First 1
        
        if ($installedModule) {
            Write-Host "  Already installed: $module (v$($installedModule.Version))" -ForegroundColor Green
        } else {
            Write-Host "  Installing: $module..." -ForegroundColor Yellow
            Install-Module -Name $module -Scope $Scope -Force -AllowClobber
            Write-Host "  Installed: $module" -ForegroundColor Green
        }
    } catch {
        Write-Error "Failed to install $module`: $($_.Exception.Message)"
    }
}

Write-Host "`nVerifying installation..." -ForegroundColor Yellow

$allInstalled = $true
foreach ($module in $requiredModules) {
    $installedModule = Get-Module -ListAvailable -Name $module
    if ($installedModule) {
        Write-Host "  [OK] $module" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] $module (NOT INSTALLED)" -ForegroundColor Red
        $allInstalled = $false
    }
}

if ($allInstalled) {
    Write-Host "`nSetup completed successfully!" -ForegroundColor Green
    Write-Host "You can now run: .\Find-AnonymousLinks.ps1" -ForegroundColor Cyan
} else {
    Write-Host "`nSetup completed with errors. Please check the installation manually." -ForegroundColor Red
}

Write-Host "`nAdditional Information:" -ForegroundColor Gray
Write-Host "- Make sure you have appropriate permissions in your Microsoft 365 tenant" -ForegroundColor Gray
Write-Host "- Global Administrator or SharePoint Administrator rights are recommended" -ForegroundColor Gray
Write-Host "- The scanner will use modern authentication (interactive popup)" -ForegroundColor Gray
