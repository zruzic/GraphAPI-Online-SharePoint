#Requires -Version 5.0
<#
.SYNOPSIS
Example: File Management Operations using GraphAPI-Client module

This script demonstrates common file operations:
- Create folders
- Upload files
- Download files
- Rename/Copy/Move files
- Search for files

.PARAMETER ClientId
Azure AD Application ID

.PARAMETER ClientSecret
Azure AD Application Secret

.PARAMETER TenantId
Azure AD Tenant ID

.PARAMETER TenantName
Tenant name (e.g., 'contoso')

.PARAMETER SiteName
SharePoint site name

.EXAMPLE
.\example-files-operations.ps1 -ClientId "xxx" -ClientSecret "yyy" `
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
    Write-Host "=== File Management Operations ===" -ForegroundColor Cyan

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

    # Create a folder
    Write-Host "`n[*] Creating folder 'TestFolder'..." -ForegroundColor Yellow
    $FolderResponse = New-GraphFolder -FolderName "TestFolder"
    $FolderId = $FolderResponse.id
    Write-Host "[+] Folder created: $($FolderResponse.name) (ID: $FolderId)" -ForegroundColor Green

    # Get folder contents
    Write-Host "`n[*] Getting root folder contents..." -ForegroundColor Yellow
    $Contents = Get-GraphFolderContents -FolderId "root"
    Write-Host "[+] Found $($Contents.Count) items in root folder" -ForegroundColor Green
    $Contents | Select-Object -First 5 | ForEach-Object {
        Write-Host "   - $($_.name) (ID: $($_.id))"
    }

    # Search for files
    Write-Host "`n[*] Searching for files with 'test' in name..." -ForegroundColor Yellow
    $SearchResults = Search-GraphFiles -SearchTerm "test"
    Write-Host "[+] Found $($SearchResults.Count) results" -ForegroundColor Green
    $SearchResults | Select-Object -First 5 | ForEach-Object {
        Write-Host "   - $($_.name)"
    }

    # Example: Create and upload a file
    Write-Host "`n[*] Creating sample file..." -ForegroundColor Yellow
    $SampleFile = Join-Path $env:TEMP "sample.txt"
    "This is a sample file for testing Graph API operations." | Set-Content -Path $SampleFile
    Write-Host "[+] Sample file created at $SampleFile"

    Write-Host "[*] Uploading file..." -ForegroundColor Yellow
    $UploadResponse = Add-GraphFile -FilePath $SampleFile -FolderId $FolderId
    $FileId = $UploadResponse.id
    Write-Host "[+] File uploaded: $($UploadResponse.name) (ID: $FileId)" -ForegroundColor Green

    # Get file metadata
    Write-Host "`n[*] Getting file metadata..." -ForegroundColor Yellow
    $Metadata = Get-GraphFileMetadata -FileId $FileId
    Write-Host "[+] File size: $($Metadata.size) bytes" -ForegroundColor Green
    Write-Host "[+] Created: $($Metadata.createdDateTime)" -ForegroundColor Green

    # Rename file
    Write-Host "`n[*] Renaming file to 'sample_renamed.txt'..." -ForegroundColor Yellow
    $RenameResponse = Rename-GraphFile -FileId $FileId -NewName "sample_renamed.txt"
    Write-Host "[+] File renamed: $($RenameResponse.name)" -ForegroundColor Green

    # Copy file
    Write-Host "`n[*] Copying file..." -ForegroundColor Yellow
    $CopyResponse = Copy-GraphFile -FileId $FileId -NewName "sample_copy.txt" -DestinationId $FolderId
    Write-Host "[+] File copy initiated" -ForegroundColor Green

    # Download file
    Write-Host "`n[*] Downloading file..." -ForegroundColor Yellow
    $DownloadPath = Join-Path $env:TEMP "downloaded_sample.txt"
    Get-GraphFile -FileId $FileId -OutputPath $DownloadPath
    Write-Host "[+] File downloaded to: $DownloadPath" -ForegroundColor Green

    # Move file to root
    Write-Host "`n[*] Moving file to root..." -ForegroundColor Yellow
    $MoveResponse = Move-GraphFile -FileId $FileId -DestinationFolderId "root"
    Write-Host "[+] File moved to root folder" -ForegroundColor Green

    # Cleanup local sample file
    Remove-Item -Path $SampleFile -Force
    Remove-Item -Path $DownloadPath -Force

    Write-Host "`n[✓] File operations example completed successfully!" -ForegroundColor Green

}
catch {
    Write-Error "[!] Error: $_"
    exit 1
}
