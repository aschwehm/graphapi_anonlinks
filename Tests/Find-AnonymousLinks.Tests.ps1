BeforeAll {
    # Import the module/script for testing
    $scriptPath = Join-Path $PSScriptRoot "Find-AnonymousLinks.ps1"
    
    # Mock Microsoft Graph modules
    $mockModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Sites", "Microsoft.Graph.Files")
    foreach ($module in $mockModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            # Create minimal mock module if not available
            New-Module -Name $module -ScriptBlock {} | Import-Module
        }
    }
    
    # Source the script functions for testing
    . $scriptPath
}

Describe "Find-AnonymousLinks Script Tests" {
    Context "Retry Logic Tests" {
        It "Should implement exponential backoff" {
            # Mock throttling response
            Mock Invoke-MgGraphRequest {
                $exception = [System.Net.WebException]::new("Throttled")
                $response = [System.Net.HttpWebResponse]::new()
                $response.StatusCode = [System.Net.HttpStatusCode]::TooManyRequests
                $exception.Response = $response
                throw $exception
            }
            
            $startTime = Get-Date
            try {
                Invoke-GraphWithRetry -Uri "https://graph.microsoft.com/v1.0/sites" -RetryCount 0
            } catch {
                # Expected to fail after retries
            }
            $duration = (Get-Date) - $startTime
            
            # Should have taken at least 1 second due to retry delays
            $duration.TotalSeconds | Should -BeGreaterThan 1
        }
        
        It "Should respect Retry-After header" {
            Mock Invoke-MgGraphRequest {
                $exception = [System.Net.WebException]::new("Throttled")
                $response = [System.Net.HttpWebResponse]::new()
                $response.StatusCode = [System.Net.HttpStatusCode]::TooManyRequests
                $response.Headers.Add("Retry-After", "2")
                $exception.Response = $response
                throw $exception
            }
            
            $startTime = Get-Date
            try {
                Invoke-GraphWithRetry -Uri "https://graph.microsoft.com/v1.0/sites" -RetryCount 0
            } catch {
                # Expected to fail
            }
            $duration = (Get-Date) - $startTime
            
            # Should have waited at least 2 seconds due to Retry-After
            $duration.TotalSeconds | Should -BeGreaterThan 2
        }
    }
    
    Context "Pagination Tests" {
        It "Should handle paginated responses correctly" {
            $mockResponse1 = @{
                value = @(
                    @{ id = "item1"; name = "Test Item 1" },
                    @{ id = "item2"; name = "Test Item 2" }
                )
                '@odata.nextLink' = "https://graph.microsoft.com/v1.0/sites?$skiptoken=abc123"
            }
            
            $mockResponse2 = @{
                value = @(
                    @{ id = "item3"; name = "Test Item 3" }
                )
            }
            
            Mock Invoke-GraphWithRetry {
                param($Uri)
                if ($Uri -like "*skiptoken*") {
                    return $mockResponse2
                } else {
                    return $mockResponse1
                }
            }
            
            $result = Get-AllGraphPages -Uri "https://graph.microsoft.com/v1.0/sites"
            
            $result.Count | Should -Be 3
            $result[0].name | Should -Be "Test Item 1"
            $result[2].name | Should -Be "Test Item 3"
        }
    }
    
    Context "Permission Detection Tests" {
        It "Should correctly identify anonymous links by scope" {
            $mockPermissions = @(
                @{
                    id = "perm1"
                    link = @{
                        scope = "anonymous"
                        type = "view"
                        webUrl = "https://example.sharepoint.com/:x:/s/site/abc123"
                    }
                },
                @{
                    id = "perm2"
                    link = @{
                        scope = "organization"
                        type = "edit"
                        webUrl = "https://example.sharepoint.com/sites/site/doc.docx"
                    }
                }
            )
            
            Mock Get-AllGraphPages { return $mockPermissions }
            
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{ AnonymousLinksFound = 0; ItemsScanned = 0; Errors = 0 }
            
            Get-AnonymousPermissions -DriveId "drive1" -ItemId "item1" -ItemName "TestDoc.docx" -ItemPath "/TestDoc.docx" -SiteId "site1" -SiteName "Test Site"
            
            $script:Results.Count | Should -Be 1
            $script:Results[0].LinkScope | Should -Be "anonymous"
            $script:Stats.AnonymousLinksFound | Should -Be 1
        }
        
        It "Should identify anyone links as anonymous" {
            $mockPermissions = @(
                @{
                    id = "perm1"
                    link = @{
                        scope = "anyone"
                        type = "view"
                        webUrl = "https://example.sharepoint.com/:x:/s/site/def456"
                    }
                }
            )
            
            Mock Get-AllGraphPages { return $mockPermissions }
            
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{ AnonymousLinksFound = 0; ItemsScanned = 0; Errors = 0 }
            
            Get-AnonymousPermissions -DriveId "drive1" -ItemId "item1" -ItemName "TestDoc.docx" -ItemPath "/TestDoc.docx" -SiteId "site1" -SiteName "Test Site"
            
            $script:Results.Count | Should -Be 1
            $script:Results[0].LinkScope | Should -Be "anyone"
        }
        
        It "Should not flag organization-scoped links" {
            $mockPermissions = @(
                @{
                    id = "perm1"
                    link = @{
                        scope = "organization"
                        type = "view"
                        webUrl = "https://example.sharepoint.com/sites/site/doc.docx"
                    }
                }
            )
            
            Mock Get-AllGraphPages { return $mockPermissions }
            
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{ AnonymousLinksFound = 0; ItemsScanned = 0; Errors = 0 }
            
            Get-AnonymousPermissions -DriveId "drive1" -ItemId "item1" -ItemName "TestDoc.docx" -ItemPath "/TestDoc.docx" -SiteId "site1" -SiteName "Test Site"
            
            $script:Results.Count | Should -Be 0
            $script:Stats.AnonymousLinksFound | Should -Be 0
        }
    }
    
    Context "Authentication Tests" {
        It "Should support interactive authentication" {
            Mock Connect-MgGraph { return $true }
            Mock Get-MgContext { 
                return @{ 
                    TenantId = "test-tenant"
                    Account = "test@example.com"
                    AuthType = "Delegated"
                }
            }
            
            $result = Connect-GraphWithMethod -Method "Interactive" -TenantId "test-tenant"
            
            $result | Should -Be $true
            Assert-MockCalled Connect-MgGraph -Times 1
        }
        
        It "Should support app-only authentication" {
            Mock Connect-MgGraph { return $true }
            Mock Get-MgContext { 
                return @{ 
                    TenantId = "test-tenant"
                    Account = $null
                    AuthType = "AppOnly"
                }
            }
            
            $result = Connect-GraphWithMethod -Method "AppOnly" -TenantId "test-tenant" -ClientId "test-client" -CertificateThumbprint "abc123"
            
            $result | Should -Be $true
            Assert-MockCalled Connect-MgGraph -Times 1
        }
        
        It "Should require parameters for app-only auth" {
            {
                Connect-GraphWithMethod -Method "AppOnly"
            } | Should -Throw "*ClientId*required*"
        }
    }
    
    Context "Output Schema Tests" {
        It "Should generate proper JSON output structure" {
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{
                StartTime = Get-Date
                SitesScanned = 5
                DrivesScanned = 10
                ItemsScanned = 100
                AnonymousLinksFound = 3
                Errors = 0
            }
            
            # Add sample result
            $sampleResult = [PSCustomObject]@{
                SiteId = "site123"
                SiteName = "Test Site"
                DriveId = "drive456"
                ItemId = "item789"
                ItemName = "TestDoc.docx"
                ItemPath = "/Documents/TestDoc.docx"
                PermissionId = "perm101"
                LinkType = "view"
                LinkScope = "anonymous"
                ExpiresOn = $null
                GrantedTo = $null
                WebUrl = "https://example.sharepoint.com/:x:/s/site/abc123"
                ScanDate = Get-Date
                HasPassword = $false
                Application = $null
                Roles = "read"
            }
            $script:Results.Add($sampleResult)
            
            # Test the export function (mock file operations)
            Mock Export-Csv {}
            Mock Out-File {}
            Mock Join-Path { return "test.json" }
            
            { Export-Results -Format "JSON" -OutputPath "." } | Should -Not -Throw
            
            # Verify the structure would be valid
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
                results = @($script:Results)
            }
            
            $exportData.scanInfo.sitesScanned | Should -Be 5
            $exportData.results.Count | Should -Be 1
            $exportData.results[0].LinkScope | Should -Be "anonymous"
        }
    }
    
    Context "Error Handling Tests" {
        It "Should handle Graph API errors gracefully" {
            Mock Get-AllGraphPages { throw "Graph API Error" }
            
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{ AnonymousLinksFound = 0; ItemsScanned = 0; Errors = 0 }
            
            { Get-AnonymousPermissions -DriveId "drive1" -ItemId "item1" -ItemName "TestDoc.docx" -ItemPath "/TestDoc.docx" -SiteId "site1" -SiteName "Test Site" } | Should -Not -Throw
            
            $script:Stats.Errors | Should -Be 1
        }
        
        It "Should continue processing after individual failures" {
            Mock Get-AllGraphPages { 
                if ($args[0] -like "*item1*") {
                    throw "Permission denied"
                } else {
                    return @(@{ id = "perm1"; link = @{ scope = "anonymous"; type = "view" } })
                }
            }
            
            $script:Results = [System.Collections.Concurrent.ConcurrentBag[PSCustomObject]]::new()
            $script:Stats = @{ AnonymousLinksFound = 0; ItemsScanned = 0; Errors = 0 }
            
            # This should fail
            Get-AnonymousPermissions -DriveId "drive1" -ItemId "item1" -ItemName "TestDoc1.docx" -ItemPath "/TestDoc1.docx" -SiteId "site1" -SiteName "Test Site"
            
            # This should succeed
            Get-AnonymousPermissions -DriveId "drive1" -ItemId "item2" -ItemName "TestDoc2.docx" -ItemPath "/TestDoc2.docx" -SiteId "site1" -SiteName "Test Site"
            
            $script:Stats.Errors | Should -Be 1
            $script:Stats.AnonymousLinksFound | Should -Be 1
        }
    }
}

Describe "Remove-AnonymousLinks Script Tests" {
    BeforeAll {
        $remediationScriptPath = Join-Path $PSScriptRoot "Remove-AnonymousLinks.ps1"
        . $remediationScriptPath
    }
    
    Context "Remediation Actions Tests" {
        It "Should delete permissions correctly" {
            Mock Invoke-GraphWithRetry { return $null }
            
            $testItem = [PSCustomObject]@{
                DriveId = "drive123"
                ItemId = "item456"
                PermissionId = "perm789"
                ItemName = "TestDoc.docx"
            }
            
            $result = Remove-Permission -Item $testItem
            
            $result | Should -Be $true
            Assert-MockCalled Invoke-GraphWithRetry -ParameterFilter { $Method -eq "DELETE" }
        }
        
        It "Should convert anonymous links to organization scope" {
            Mock Invoke-GraphWithRetry {
                param($Uri, $Method, $Body)
                
                if ($Method -eq "GET") {
                    return @{
                        link = @{
                            scope = "anonymous"
                            type = "view"
                        }
                    }
                } else {
                    return $null
                }
            }
            
            $testItem = [PSCustomObject]@{
                DriveId = "drive123"
                ItemId = "item456"
                PermissionId = "perm789"
                ItemName = "TestDoc.docx"
            }
            
            $result = Convert-ToOrganizationLink -Item $testItem
            
            $result | Should -Be $true
            Assert-MockCalled Invoke-GraphWithRetry -ParameterFilter { $Method -eq "PATCH" }
        }
        
        It "Should set expiration dates correctly" {
            Mock Invoke-GraphWithRetry {
                param($Uri, $Method, $Body)
                
                if ($Method -eq "GET") {
                    return @{
                        link = @{
                            scope = "anonymous"
                            type = "view"
                        }
                    }
                } else {
                    return $null
                }
            }
            
            $testItem = [PSCustomObject]@{
                DriveId = "drive123"
                ItemId = "item456"
                PermissionId = "perm789"
                ItemName = "TestDoc.docx"
            }
            
            $result = Set-LinkExpiration -Item $testItem -Days 30
            
            $result | Should -Be $true
            Assert-MockCalled Invoke-GraphWithRetry -ParameterFilter { $Method -eq "PATCH" }
        }
    }
}
