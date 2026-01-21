#Requires -Version 5.0
<#
.SYNOPSIS
Example: SharePoint Pages Management using GraphAPI-Client module

This script demonstrates page operations:
- Create pages
- Get page details
- Update page properties
- Publish pages
- Delete pages

.PARAMETER ClientId
Azure AD Application ID

.PARAMETER ClientSecret
Azure AD Application Secret

.PARAMETER TenantId
Azure AD Tenant ID

.PARAMETER TenantName
Tenant name

.PARAMETER SiteName
SharePoint site name

.EXAMPLE
.\example-pages-operations.ps1 -ClientId "xxx" -ClientSecret "yyy" `
    -TenantId "zzz" -TenantName "contoso" -SiteName "TeamSite"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$TenantName,

    [Parameter(Mandatory = $true)]
    [string]$SiteName
)

# Import the GraphAPI module
$ModulePath = Join-Path $PSScriptRoot "GraphAPI-Client.psm1"
if (-not (Test-Path $ModulePath)) {
    Write-Error "GraphAPI-Client.psm1 not found at $ModulePath"
    exit 1
}

Import-Module $ModulePath -Force

try {
    Write-Host "=== Pages Management Operations ===" -ForegroundColor Cyan

    # Authenticate
    Write-Host "`n[*] Authenticating..." -ForegroundColor Yellow
    Get-GraphAccessToken -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId

    # Get Site ID
    Write-Host "[*] Getting Site ID for site: $SiteName" -ForegroundColor Yellow
    $SiteId = Get-GraphSiteId -SiteName $SiteName -TenantName $TenantName
    Write-Host "[+] Site ID: $SiteId" -ForegroundColor Green

    # Get all pages
    Write-Host "`n[*] Retrieving all pages..." -ForegroundColor Yellow
    $Pages = Get-GraphPages
    Write-Host "[+] Found $($Pages.Count) pages" -ForegroundColor Green
    $Pages | Select-Object -First 5 | ForEach-Object {
        Write-Host "   - $($_.name) (Title: $($_.title))"
    }

    # Create a new page
    Write-Host "`n[*] Creating new page 'DemoPage.aspx'..." -ForegroundColor Yellow
    $PageResponse = New-GraphPage -Name "DemoPage.aspx" -Title "Demo Page Title"
    $PageId = $PageResponse.id
    Write-Host "[+] Page created: $($PageResponse.name)" -ForegroundColor Green
    Write-Host "[+] ID: $PageId" -ForegroundColor Green
    Write-Host "[+] URL: $($PageResponse.webUrl)" -ForegroundColor Green

    # Get page details
    Write-Host "`n[*] Getting page details..." -ForegroundColor Yellow
    $PageDetails = Get-GraphPageDetails -PageId $PageId
    Write-Host "[+] Page Title: $($PageDetails.title)" -ForegroundColor Green
    Write-Host "[+] Description: $($PageDetails.description)" -ForegroundColor Green
    Write-Host "[+] Status: $($PageDetails.publishingState)" -ForegroundColor Green

    # Update page
    Write-Host "`n[*] Updating page title and description..." -ForegroundColor Yellow
    $UpdateResponse = Set-GraphPage -PageId $PageId `
        -Title "Updated Demo Page" `
        -Description "This is an updated demo page with more information"
    Write-Host "[+] Page updated: $($UpdateResponse.title)" -ForegroundColor Green

    # Add web part to page
    Write-Host "`n[*] Adding web part to page..." -ForegroundColor Yellow
    $WebPartData = @{
        id           = "webpartid"
        instanceId   = "00000000-0000-0000-0000-000000000000"
        title        = "Welcome Text"
        dataVersion  = "1.0"
        properties   = @{
            text = "Welcome to this page"
        }
        serverProcessedContent = @{
            htmlStrings         = @{}
            searchablePlainTexts = @("Welcome to this page")
            imageSources        = @()
            links               = @()
        }
    }

    try {
        $WebPartResponse = Add-GraphWebPart -PageId $PageId -WebPartData $WebPartData
        Write-Host "[+] Web part added successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Web part addition error (may be expected): $_" -ForegroundColor Yellow
    }

    # Publish the page
    Write-Host "`n[*] Publishing page..." -ForegroundColor Yellow
    try {
        Publish-GraphPage -PageId $PageId
        Write-Host "[+] Page published successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Publishing error: $_" -ForegroundColor Yellow
    }

    # Get updated page details
    Write-Host "`n[*] Getting final page details..." -ForegroundColor Yellow
    $FinalDetails = Get-GraphPageDetails -PageId $PageId
    Write-Host "[+] Final Title: $($FinalDetails.title)" -ForegroundColor Green
    Write-Host "[+] Status: $($FinalDetails.publishingState)" -ForegroundColor Green

    # Note: Uncomment to delete the page
    # Write-Host "`n[*] Deleting page..." -ForegroundColor Yellow
    # Remove-GraphPage -PageId $PageId
    # Write-Host "[+] Page deleted" -ForegroundColor Green

    Write-Host "`n[✓] Pages management example completed successfully!" -ForegroundColor Green

}
catch {
    Write-Error "[!] Error: $_"
    exit 1
}
