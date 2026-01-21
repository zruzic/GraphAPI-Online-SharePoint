#Requires -Version 5.0
<#
.SYNOPSIS
Example: Security & Sharing Operations using GraphAPI-Client module

This script demonstrates security operations:
- Get item permissions
- Share files with users
- Create sharing links
- Update permission roles
- Delete permissions

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

.PARAMETER FileId
File ID to share

.PARAMETER UserEmail
User email to share with

.EXAMPLE
.\example-security-operations.ps1 -ClientId "xxx" -ClientSecret "yyy" `
    -TenantId "zzz" -TenantName "contoso" -SiteName "TeamSite" `
    -FileId "file_id" -UserEmail "user@example.com"
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
    [string]$SiteName,

    [Parameter(Mandatory = $false)]
    [string]$FileId = "file_id_to_share",

    [Parameter(Mandatory = $false)]
    [string]$UserEmail = "user@example.com"
)

# Import the GraphAPI module
$ModulePath = Join-Path $PSScriptRoot "GraphAPI-Client.psm1"
if (-not (Test-Path $ModulePath)) {
    Write-Error "GraphAPI-Client.psm1 not found at $ModulePath"
    exit 1
}

Import-Module $ModulePath -Force

try {
    Write-Host "=== Security & Sharing Operations ===" -ForegroundColor Cyan

    # Authenticate
    Write-Host "`n[*] Authenticating..." -ForegroundColor Yellow
    Get-GraphAccessToken -ClientId $ClientId -ClientSecret $ClientSecret -TenantId $TenantId

    # Get Site ID
    Write-Host "[*] Getting Site ID for site: $SiteName" -ForegroundColor Yellow
    $SiteId = Get-GraphSiteId -SiteName $SiteName -TenantName $TenantName
    Write-Host "[+] Site ID: $SiteId" -ForegroundColor Green

    # Get Drive ID
    Write-Host "[*] Getting Drive ID..." -ForegroundColor Yellow
    $DriveId = Get-GraphDriveId
    Write-Host "[+] Drive ID: $DriveId" -ForegroundColor Green

    # Get file metadata first
    Write-Host "`n[*] Getting file metadata for $FileId..." -ForegroundColor Yellow
    try {
        $FileMetadata = Get-GraphFileMetadata -FileId $FileId
        Write-Host "[+] File: $($FileMetadata.name)" -ForegroundColor Green
    }
    catch {
        Write-Host "[*] File ID may not exist (using for example): $_" -ForegroundColor Yellow
    }

    # Get current permissions
    Write-Host "`n[*] Getting current permissions for file..." -ForegroundColor Yellow
    try {
        $Permissions = Get-GraphItemPermissions -FileId $FileId
        Write-Host "[+] Found $($Permissions.Count) permissions:" -ForegroundColor Green
        foreach ($Perm in $Permissions) {
            Write-Host "   - Permission ID: $($Perm.id)"
            if ($Perm.roles) {
                Write-Host "     Roles: $($Perm.roles -join ', ')"
            }
        }
    }
    catch {
        Write-Host "[*] Could not retrieve permissions (file may not exist): $_" -ForegroundColor Yellow
    }

    # Share with user (Edit access)
    Write-Host "`n[*] Sharing file with user (Edit access)..." -ForegroundColor Yellow
    try {
        $ShareResponse = Add-GraphItemShare -FileId $FileId -Email $UserEmail -Role "edit"
        Write-Host "[+] Share invitation sent to $UserEmail" -ForegroundColor Green
        if ($ShareResponse.value) {
            foreach ($Grant in $ShareResponse.value) {
                Write-Host "   Permission ID: $($Grant.id)"
                Write-Host "   Roles: $($Grant.roles -join ', ')"
            }
        }
    }
    catch {
        Write-Host "[!] Share operation error (file may not exist): $_" -ForegroundColor Yellow
    }

    # Share with user (Read-only access)
    Write-Host "`n[*] Sharing file with user (Read-only access)..." -ForegroundColor Yellow
    try {
        $ShareReadonly = Add-GraphItemShare -FileId $FileId -Email $UserEmail -Role "read"
        Write-Host "[+] Read-only share invitation sent" -ForegroundColor Green
    }
    catch {
        Write-Host "[!] Read-only share error: $_" -ForegroundColor Yellow
    }

    # Create anonymous sharing link (View)
    Write-Host "`n[*] Creating anonymous view-only sharing link..." -ForegroundColor Yellow
    try {
        $LinkView = New-GraphSharingLink -FileId $FileId -LinkType "view" -Scope "anonymous"
        if ($LinkView.link) {
            Write-Host "[+] View link created:" -ForegroundColor Green
            Write-Host "   URL: $($LinkView.link.webUrl)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "[!] View link creation error: $_" -ForegroundColor Yellow
    }

    # Create anonymous sharing link (Edit)
    Write-Host "`n[*] Creating anonymous edit sharing link..." -ForegroundColor Yellow
    try {
        $LinkEdit = New-GraphSharingLink -FileId $FileId -LinkType "edit" -Scope "anonymous"
        if ($LinkEdit.link) {
            Write-Host "[+] Edit link created:" -ForegroundColor Green
            Write-Host "   URL: $($LinkEdit.link.webUrl)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "[!] Edit link creation error: $_" -ForegroundColor Yellow
    }

    # Update permission role
    Write-Host "`n[*] Updating permission role..." -ForegroundColor Yellow
    if ($Permissions -and $Permissions.Count -gt 0) {
        $PermId = $Permissions[0].id
        try {
            $UpdatedPerm = Set-GraphItemShareRole -FileId $FileId -PermissionId $PermId -Role "read"
            Write-Host "[+] Permission updated to read-only" -ForegroundColor Green
            Write-Host "   New roles: $($UpdatedPerm.roles -join ', ')" -ForegroundColor Green
        }
        catch {
            Write-Host "[!] Update permission error: $_" -ForegroundColor Yellow
        }
    }

    # Delete permission
    Write-Host "`n[*] Deleting permission..." -ForegroundColor Yellow
    if ($Permissions -and $Permissions.Count -gt 1) {
        $PermId = $Permissions[1].id
        try {
            Remove-GraphItemShare -FileId $FileId -PermissionId $PermId
            Write-Host "[+] Permission deleted" -ForegroundColor Green
        }
        catch {
            Write-Host "[!] Delete permission error: $_" -ForegroundColor Yellow
        }
    }

    Write-Host "`n[✓] Security & Sharing example completed!" -ForegroundColor Green

}
catch {
    Write-Error "[!] Error: $_"
    exit 1
}
