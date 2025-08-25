<#
.SYNOPSIS
    Setup script for SharePoint Anonymous Links Scanner v3.0

.DESCRIPTION
    This script installs the required Microsoft Graph PowerShell modules,
    PSScriptAnalyzer, Pester, and sets up the environment for running 
    the production-ready anonymous links scanner and remediation tools.

.PARAMETER Scope
    Installation scope - 'CurrentUser' (default) or 'AllUsers'

.PARAMETER IncludeDevTools
    Install development tools (Pester, PSScriptAnalyzer) for testing

.PARAMETER Force
    Force reinstall modules even if they exist

.EXAMPLE
    .\Setup.ps1
    Installs modules for current user

.EXAMPLE
    .\Setup.ps1 -Scope AllUsers -IncludeDevTools
    Installs modules for all users with dev tools (requires admin rights)

.EXAMPLE
    .\Setup.ps1 -Force
    Force reinstall all modules
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("CurrentUser", "AllUsers")]
    [string]$Scope = "CurrentUser",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDevTools,
    
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Write-Host "SharePoint Anonymous Links Scanner v3.0 - Setup" -ForegroundColor Green
Write-Host "===============================================" -ForegroundColor Green

# Required modules for core functionality
$requiredModules = @(
    @{ Name = "Microsoft.Graph.Authentication"; MinVersion = "1.12.0" },
    @{ Name = "Microsoft.Graph.Sites"; MinVersion = "1.12.0" },
    @{ Name = "Microsoft.Graph.Files"; MinVersion = "1.12.0" }
)

# Development tools
$devModules = @(
    @{ Name = "PSScriptAnalyzer"; MinVersion = "1.20.0" },
    @{ Name = "Pester"; MinVersion = "5.3.0" }
)

Write-Host "Installation scope: $Scope" -ForegroundColor Cyan
Write-Host "Include dev tools: $($IncludeDevTools.IsPresent)" -ForegroundColor Cyan
Write-Host "Force reinstall: $($Force.IsPresent)" -ForegroundColor Cyan
Write-Host ""

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion
Write-Host "PowerShell version: $($psVersion.ToString())" -ForegroundColor Gray

if ($psVersion.Major -lt 5) {
    Write-Error "PowerShell 5.0 or later is required. Current version: $($psVersion.ToString())"
    exit 1
}

if ($psVersion.Major -eq 5 -and $psVersion.Minor -eq 0) {
    Write-Warning "PowerShell 5.0 detected. Consider upgrading to PowerShell 5.1 or 7.x for better performance."
}

# Set execution policy if needed
try {
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-Host "Current execution policy (CurrentUser): $currentPolicy" -ForegroundColor Gray
    
    if ($currentPolicy -eq "Restricted" -or $currentPolicy -eq "Undefined") {
        Write-Host "Setting execution policy to RemoteSigned for CurrentUser..." -ForegroundColor Yellow
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
        Write-Host "Execution policy updated" -ForegroundColor Green
    }
} catch {
    Write-Warning "Could not set execution policy: $($_.Exception.Message)"
}

# Set PSGallery as trusted if needed
try {
    $gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($gallery -and $gallery.InstallationPolicy -ne "Trusted") {
        Write-Host "Setting PSGallery as trusted repository..." -ForegroundColor Yellow
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        Write-Host "PSGallery set as trusted" -ForegroundColor Green
    }
} catch {
    Write-Warning "Could not set PSGallery as trusted: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Installing required modules..." -ForegroundColor Cyan

# Function to install or update a module
function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [string]$MinVersion = $null,
        [string]$InstallScope,
        [bool]$ForceInstall = $false
    )
    
    try {
        $existingModule = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        
        if ($existingModule -and -not $ForceInstall) {
            if ($MinVersion) {
                $minVersionObj = [Version]$MinVersion
                if ($existingModule.Version -ge $minVersionObj) {
                    Write-Host "  ✓ $ModuleName ($($existingModule.Version)) - Already installed" -ForegroundColor Green
                    return $true
                } else {
                    Write-Host "  ↑ $ModuleName - Updating from $($existingModule.Version) to latest..." -ForegroundColor Yellow
                }
            } else {
                Write-Host "  ✓ $ModuleName ($($existingModule.Version)) - Already installed" -ForegroundColor Green
                return $true
            }
        } else {
            Write-Host "  ⬇ $ModuleName - Installing..." -ForegroundColor Yellow
        }
        
        $installParams = @{
            Name = $ModuleName
            Scope = $InstallScope
            Force = $ForceInstall
            AllowClobber = $true
        }
        
        if ($MinVersion) {
            $installParams.MinimumVersion = $MinVersion
        }
        
        Install-Module @installParams
        
        # Verify installation
        $newModule = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($newModule) {
            Write-Host "  ✓ $ModuleName ($($newModule.Version)) - Installed successfully" -ForegroundColor Green
            return $true
        } else {
            Write-Host "  ✗ $ModuleName - Installation verification failed" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "  ✗ $ModuleName - Installation failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Install core modules
$installationSuccess = $true
foreach ($module in $requiredModules) {
    $success = Install-RequiredModule -ModuleName $module.Name -MinVersion $module.MinVersion -InstallScope $Scope -ForceInstall $Force.IsPresent
    if (-not $success) {
        $installationSuccess = $false
    }
}

# Install development tools if requested
if ($IncludeDevTools) {
    Write-Host ""
    Write-Host "Installing development tools..." -ForegroundColor Cyan
    
    foreach ($module in $devModules) {
        $success = Install-RequiredModule -ModuleName $module.Name -MinVersion $module.MinVersion -InstallScope $Scope -ForceInstall $Force.IsPresent
        if (-not $success) {
            $installationSuccess = $false
        }
    }
}

Write-Host ""

if ($installationSuccess) {
    Write-Host "✓ Setup completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "1. Run the scanner: .\Find-AnonymousLinks.ps1" -ForegroundColor White
    Write-Host "2. Review results in generated CSV/JSON files" -ForegroundColor White
    Write-Host "3. Use .\Remove-AnonymousLinks.ps1 for remediation" -ForegroundColor White
    
    if ($IncludeDevTools) {
        Write-Host ""
        Write-Host "Development tools installed:" -ForegroundColor Cyan
        Write-Host "- Run tests: Invoke-Pester -Path .\Tests\" -ForegroundColor White
        Write-Host "- Code analysis: Invoke-ScriptAnalyzer -Path .\Find-AnonymousLinks.ps1" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "For help and examples, see README.md" -ForegroundColor Gray
} else {
    Write-Host "✗ Setup encountered errors. Please review the output above." -ForegroundColor Red
    Write-Host ""
    Write-Host "Manual installation commands:" -ForegroundColor Yellow
    foreach ($module in $requiredModules) {
        Write-Host "Install-Module -Name $($module.Name) -Scope $Scope -Force" -ForegroundColor Gray
    }
    exit 1
}

# Test module loading
Write-Host ""
Write-Host "Testing module imports..." -ForegroundColor Cyan

$testSuccess = $true
foreach ($module in $requiredModules) {
    try {
        Import-Module -Name $module.Name -Force -ErrorAction Stop
        Write-Host "  ✓ $($module.Name) - Import successful" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ $($module.Name) - Import failed: $($_.Exception.Message)" -ForegroundColor Red
        $testSuccess = $false
    }
}

if ($testSuccess) {
    Write-Host ""
    Write-Host "🎉 All modules imported successfully! Ready to scan for anonymous links." -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "⚠️  Some modules failed to import. You may need to restart PowerShell or check for conflicts." -ForegroundColor Yellow
}
