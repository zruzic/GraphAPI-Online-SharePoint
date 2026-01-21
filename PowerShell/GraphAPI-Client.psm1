#Requires -Version 5.0
<#
.SYNOPSIS
GraphAPI Client Module for Microsoft Graph API - SharePoint Operations

.DESCRIPTION
PowerShell module providing functions for authentication, file management,
document sets, security, and pages operations on SharePoint via Microsoft Graph API.

.AUTHOR
SharePoint Graph Documentation

.VERSION
1.0.0
#>

$script:GraphAPIBaseUrl = "https://graph.microsoft.com/v1.0"
$script:AccessToken = $null
$script:SiteId = $null
$script:DriveId = $null

<#
.SYNOPSIS
Get access token from Azure AD using client credentials flow.

.PARAMETER ClientId
Azure AD Application ID

.PARAMETER ClientSecret
Azure AD Application Secret

.PARAMETER TenantId
Azure AD Tenant ID

.EXAMPLE
$token = Get-GraphAccessToken -ClientId "xxx" -ClientSecret "yyy" -TenantId "zzz"
#>
function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    $TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $Body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = "https://graph.microsoft.com/.default"
    }

    try {
        $Response = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $Body -ErrorAction Stop
        $script:AccessToken = $Response.access_token
        Write-Host "Access token obtained successfully." -ForegroundColor Green
        return $Response.access_token
    }
    catch {
        Write-Error "Failed to get access token: $_"
        throw
    }
}

<#
.SYNOPSIS
Get Site ID from site name.

.PARAMETER SiteName
SharePoint site name

.PARAMETER TenantName
Tenant name (e.g., 'contoso')
#>
function Get-GraphSiteId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteName,

        [Parameter(Mandatory = $true)]
        [string]$TenantName
    )

    if (-not $script:AccessToken) {
        Write-Error "Access token not set. Call Get-GraphAccessToken first."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$TenantName.sharepoint.com:/sites/$SiteName"
    $Headers = @{
        "Authorization" = "Bearer $script:AccessToken"
    }

    try {
        $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers -ErrorAction Stop
        $script:SiteId = $Response.id
        Write-Host "Site ID retrieved: $script:SiteId" -ForegroundColor Green
        return $script:SiteId
    }
    catch {
        Write-Error "Failed to get site ID: $_"
        throw
    }
}

<#
.SYNOPSIS
Get Drive ID (default document library) for the site.
#>
function Get-GraphDriveId {
    [CmdletBinding()]
    param()

    if (-not $script:SiteId) {
        Write-Error "Site ID not set. Call Get-GraphSiteId first."
        throw
    }

    if (-not $script:AccessToken) {
        Write-Error "Access token not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/drives"
    $Headers = @{
        "Authorization" = "Bearer $script:AccessToken"
    }

    try {
        $Response = Invoke-RestMethod -Uri $Uri -Method Get -Headers $Headers -ErrorAction Stop
        $script:DriveId = $Response.value[0].id
        Write-Host "Drive ID retrieved: $script:DriveId" -ForegroundColor Green
        return $script:DriveId
    }
    catch {
        Write-Error "Failed to get drive ID: $_"
        throw
    }
}

<#
.SYNOPSIS
Helper function to make Graph API requests.
#>
function Invoke-GraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [string]$Method,

        [Parameter(Mandatory = $false)]
        [object]$Body,

        [Parameter(Mandatory = $false)]
        [hashtable]$Headers = @{}
    )

    if (-not $script:AccessToken) {
        Write-Error "Access token not set."
        throw
    }

    $DefaultHeaders = @{
        "Authorization" = "Bearer $script:AccessToken"
        "Content-Type"  = "application/json"
    }

    $Headers.Keys | ForEach-Object { $DefaultHeaders[$_] = $Headers[$_] }

    try {
        $Params = @{
            Uri     = $Uri
            Method  = $Method
            Headers = $DefaultHeaders
            ErrorAction = "Stop"
        }

        if ($Body) {
            $Params["Body"] = $Body | ConvertTo-Json -Depth 10
        }

        $Response = Invoke-RestMethod @Params
        return $Response
    }
    catch {
        Write-Error "Graph API request failed: $_"
        throw
    }
}

# ==================== File Management ====================

<#
.SYNOPSIS
Create a folder.
#>
function New-GraphFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName,

        [Parameter(Mandatory = $false)]
        [string]$ParentId = "root"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$ParentId/children"

    $Body = @{
        name                                = $FolderName
        folder                              = @{}
        "@microsoft.graph.conflictBehavior" = "rename"
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Delete a file.
#>
function Remove-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId"
    Invoke-GraphRequest -Uri $Uri -Method Delete
}

<#
.SYNOPSIS
Rename a file.
#>
function Rename-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$NewName
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId"

    $Body = @{
        name = $NewName
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

<#
.SYNOPSIS
Copy a file.
#>
function Copy-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$NewName,

        [Parameter(Mandatory = $false)]
        [string]$DestinationId = "root"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/copy"

    $Body = @{
        parentReference = @{
            driveId = $script:DriveId
            id      = $DestinationId
        }
        name            = $NewName
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Move a file to a different folder.
#>
function Move-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFolderId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId"

    $Body = @{
        parentReference = @{
            id = $DestinationFolderId
        }
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

<#
.SYNOPSIS
Get file metadata.
#>
function Get-GraphFileMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId"
    return Invoke-GraphRequest -Uri $Uri -Method Get
}

<#
.SYNOPSIS
Get folder contents.
#>
function Get-GraphFolderContents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FolderId/children"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Search files by name.
#>
function Search-GraphFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/root/search(q='$SearchTerm')"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Download a file.
#>
function Get-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/content"

    try {
        $Headers = @{
            "Authorization" = "Bearer $script:AccessToken"
        }
        Invoke-WebRequest -Uri $Uri -Method Get -Headers $Headers -OutFile $OutputPath -ErrorAction Stop
        Write-Host "File downloaded to: $OutputPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to download file: $_"
        throw
    }
}

<#
.SYNOPSIS
Upload a file (simple upload for files < 4MB).
#>
function Add-GraphFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [Parameter(Mandatory = $false)]
        [string]$FolderId = "root"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $FileName = Split-Path -Leaf $FilePath
    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FolderId`:/$FileName`:/content"

    try {
        $Headers = @{
            "Authorization" = "Bearer $script:AccessToken"
        }
        $FileContent = [System.IO.File]::ReadAllBytes($FilePath)
        $Response = Invoke-RestMethod -Uri $Uri -Method Put -Headers $Headers -Body $FileContent -ErrorAction Stop
        Write-Host "File uploaded successfully" -ForegroundColor Green
        return $Response
    }
    catch {
        Write-Error "Failed to upload file: $_"
        throw
    }
}

# ==================== Document Sets ====================

<#
.SYNOPSIS
Get all document sets in a list.
#>
function Get-GraphDocumentSets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items?`$filter=contentType/name eq 'Document Set'"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Get document set details.
#>
function Get-GraphDocumentSetDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$DocSetId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$DocSetId"
    return Invoke-GraphRequest -Uri $Uri -Method Get
}

<#
.SYNOPSIS
Create a document set.
#>
function New-GraphDocumentSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items"

    $Body = @{
        fields = @{
            Title         = $Title
            ContentTypeId = "0x0120D520"
        }
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Update document set properties.
#>
function Set-GraphDocumentSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$DocSetId,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$DocSetId"

    $Body = @{
        fields = @{
            Title         = $Title
            ContentTypeId = "0x0120D520"
        }
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

<#
.SYNOPSIS
Delete a document set.
#>
function Remove-GraphDocumentSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$DocSetId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$DocSetId"
    Invoke-GraphRequest -Uri $Uri -Method Delete
}

# ==================== Security & Sharing ====================

<#
.SYNOPSIS
Get item permissions.
#>
function Get-GraphItemPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/permissions"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Share item with a user.
#>
function Add-GraphItemShare {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $false)]
        [string]$Role = "edit"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/invite"

    $Body = @{
        recipients   = @(@{
                email = $Email
            })
        roles        = @($Role)
        requireSignIn = $true
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Create a sharing link.
#>
function New-GraphSharingLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $false)]
        [ValidateSet("view", "edit")]
        [string]$LinkType = "view",

        [Parameter(Mandatory = $false)]
        [string]$Scope = "anonymous"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/createLink"

    $Body = @{
        type  = $LinkType
        scope = $Scope
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Delete a permission.
#>
function Remove-GraphItemShare {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$PermissionId
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/permissions/$PermissionId"
    Invoke-GraphRequest -Uri $Uri -Method Delete
}

<#
.SYNOPSIS
Update permission role.
#>
function Set-GraphItemShareRole {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileId,

        [Parameter(Mandatory = $true)]
        [string]$PermissionId,

        [Parameter(Mandatory = $false)]
        [string]$Role = "read"
    )

    if (-not $script:DriveId) {
        Write-Error "Drive ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/drives/$script:DriveId/items/$FileId/permissions/$PermissionId"

    $Body = @{
        roles = @($Role)
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

# ==================== Pages Management ====================

<#
.SYNOPSIS
Get all pages in the site.
#>
function Get-GraphPages {
    [CmdletBinding()]
    param()

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Get page details.
#>
function Get-GraphPageDetails {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages/$PageId"
    return Invoke-GraphRequest -Uri $Uri -Method Get
}

<#
.SYNOPSIS
Create a new page.
#>
function New-GraphPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages"

    $Body = @{
        name               = $Name
        title              = $Title
        layoutWebpartId    = "3eb3e627-5144-4667-83d5-7662c6abb714"
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Update page title and description.
#>
function Set-GraphPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageId,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $false)]
        [string]$Description = ""
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages/$PageId"

    $Body = @{
        title       = $Title
        description = $Description
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

<#
.SYNOPSIS
Publish a page.
#>
function Publish-GraphPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages/$PageId/publish"
    Invoke-GraphRequest -Uri $Uri -Method Post
    Write-Host "Page published successfully" -ForegroundColor Green
}

<#
.SYNOPSIS
Delete a page.
#>
function Remove-GraphPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages/$PageId"
    Invoke-GraphRequest -Uri $Uri -Method Delete
    Write-Host "Page deleted successfully" -ForegroundColor Green
}

<#
.SYNOPSIS
Add a web part to a page.
#>
function Add-GraphWebPart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageId,

        [Parameter(Mandatory = $true)]
        [object]$WebPartData
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/pages/$PageId/webparts"

    $Body = @{
        webPartData = $WebPartData
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

# ==================== List Management ====================

<#
.SYNOPSIS
Get all lists in the site.
#>
function Get-GraphLists {
    [CmdletBinding()]
    param()

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists"
    $Response = Invoke-GraphRequest -Uri $Uri -Method Get
    return $Response.value
}

<#
.SYNOPSIS
Create a list item.
#>
function New-GraphListItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Fields
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items"

    $Body = @{
        fields = $Fields
    }

    return Invoke-GraphRequest -Uri $Uri -Method Post -Body $Body
}

<#
.SYNOPSIS
Get a list item.
#>
function Get-GraphListItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$ItemId"
    return Invoke-GraphRequest -Uri $Uri -Method Get
}

<#
.SYNOPSIS
Update a list item.
#>
function Set-GraphListItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Fields
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$ItemId"

    $Body = @{
        fields = $Fields
    }

    return Invoke-GraphRequest -Uri $Uri -Method Patch -Body $Body
}

<#
.SYNOPSIS
Delete a list item.
#>
function Remove-GraphListItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ListId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )

    if (-not $script:SiteId) {
        Write-Error "Site ID not set."
        throw
    }

    $Uri = "$script:GraphAPIBaseUrl/sites/$script:SiteId/lists/$ListId/items/$ItemId"
    Invoke-GraphRequest -Uri $Uri -Method Delete
}

Export-ModuleMember -Function Get-Graph*, New-Graph*, Set-Graph*, Remove-Graph*, Add-Graph*, Copy-Graph*, Move-Graph*, Publish-Graph*, Search-Graph*, Invoke-Graph*, Rename-Graph*
