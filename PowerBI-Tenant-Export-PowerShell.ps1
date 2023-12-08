<#
Copyright (c) Tailored Technical Solutions LLC.

MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

# Dependencies
# Install PowerShell 7.3+
# Check your vesion: $PSVersionTable.PSVersion 
# Install PowerShell Power BI command lets
# https://learn.microsoft.com/en-us/powershell/power-bi/overview?view=powerbi-ps

<# 
If you’re getting the error message ‘PowerShell script is not digitally signed,’ you will probably see that your device’s (or current user’s) execution policy is set 
 to ‘AllSigned’ or ‘Remote Signed’ (in the case of a Window server machine). If you’re not seeing the error message but can’t run the script, it’s likely because your execution 
 policies are set to ‘Restricted’ or ‘Undefined.’

DO NOT BYPASS GLOBALLY. You will make your system insecure. 
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

OR

Create your own signature
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing?view=powershell-7.3

The script has been tested against tenants with workspaces created in 2023. Older Power BI tenant object versions might produce errors. I have tried to make the script flexible.
It has been tested with Windows 11 and Mac users. Windows 11 users seem to have more success.
#>

# https://learn.microsoft.com/en-us/power-bi/admin/service-admin-auditing
# https://learn.microsoft.com/en-us/power-bi/collaborate-share/service-data-source-impact-analysis
# https://powerbi.microsoft.com/en-us/blog/announcing-scanner-api-admin-rest-apis-enhancements-to-include-dataset-tables-columns-measures-dax-expressions-and-mashup-queries/
# https://learn.microsoft.com/en-us/rest/api/power-bi/admin/workspace-info-post-workspace-info

# It is a good idea to kick off a metadata scan.

[CmdletBinding(DefaultParameterSetName = "StartDateTime")]
param
(
  [string] $OutputDirectory
)

function CreateDirectoryIfNotExists  {
  param (
    [string] $path
  )

  if (!(Test-Path $path)) {
      New-Item -Path $path -ItemType Directory
  }
}

$OutputDirectory = Read-Host -Prompt "Enter the directory"
$RetrieveDate = Get-Date

if ($OutputDirectory.Length -eq 0) {
  $OutputDirectory = 'C:\PBI-Export'
}

CreateDirectoryIfNotExists($OutputDirectory)

$DateToString = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
$FullLogPath = $OutputDirectory + "\$DateToString" + "PowerBI.log"
Start-Transcript -Path $FullLogPath

# Delete previous data run files. Keep the log files for debugging purposes.
Get-ChildItem -Path $OutputDirectory -Include *.* -File -Recurse | Where-Object Name -NotLike "*PowerBI.log" | ForEach-Object { $_.Delete() }

# Environment
$user = Connect-PowerBIServiceAccount
$user | Export-Csv "$OutputDirectory\Environment.csv"

# Workspaces
# This call will sometimes return duplicate Workspace ID records. It does return Workspace description, which can be useful.
# $WorkspacesMoreInfo = Get-PowerBIWorkspace -All -Scope Organization -Include All
# $WorkspacesMoreInfo | Export-Csv "$OutputDirectory\WorkspacesMoreInfo.csv"

$Response = Invoke-PowerBIRestMethod -Url "https://api.powerbi.com/v1.0/myorg/groups" -Body $Body -Method Get | ConvertFrom-Json 
$Response.value | Export-Csv "$OutputDirectory\Workspaces-REST.csv"

$Workspaces

# Returns all workspaces in the tenant and adds capacity information.
# Maximum 50 requests per hour, per tenant. 
function Export-WorkspacesAsAdmin {
  Write-Host "ExportWorkspacesAsAdmin"
  $WorkspaceLimit = 2000
  $Uri = 'https://api.powerbi.com/v1.0/myorg/admin/groups?$expand=reports,datasets,users,dashboards,workbooks&$top=' + $WorkspaceLimit
  $Response = Invoke-PowerBIRestMethod -Url $Uri -Method Get | ConvertFrom-Json 
  $global:Workspaces = $Response.value
  $global:Workspaces | Export-Csv "$OutputDirectory\WorkspacesAdmin-REST.csv"
} 

# Once you get the main workspaces, then kick off a scanning query.
# TODO: Finish this code.
function PostWorkspaceInfo {
  param(
    $Workspaces
  )

  Write-Host "Calling ExportWorkspaceMetadata"
  Write-Host "Start metadata scan."
  $ScanIds = [System.Collections.Generic.List[String]]::new()

  foreach($Workspace in $Workspaces) {
    if ($Workspace.Type -ne "PersonalGroup") {

      $Body = '{
        "workspaces": [
          "' + $Workspace.Id + '"
        ]
      }' | ConvertFrom-Json | ConvertTo-Json

      $Uri = "https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True"
      $Response = Invoke-PowerBIRestMethod -Url $Uri -Method Post | ConvertFrom-Json
      $ScanIds.Add($Response.Id)
      $Response | Export-Csv "$OutputDirectory\WorkspaceMetadataScan.csv" -Append
    }
  }

  $ScanIdsReady = [System.Collections.Generic.List[String]]::new()

  foreach($ScanId in $ScanIds)
  {
    $Response.status = "Not started"

    while ( $Response | Where-Object {$_.status -notmatch "Failed" -and $_.status -notmatch "Succeeded"}) {
      Start-Sleep -Seconds 1.0
      $Uri =  "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanStatus/$ScanId"
      $Response = Invoke-PowerBIRestMethod -Url $Uri -Method Get | ConvertFrom-Json
      $Response | Export-Csv "$OutputDirectory\WorkspaceMetadataScanStatus.csv" -Append
      
      if ("Succeeded" -eq $Response.status){
        $ScanIdsReady.Add($ScanId)
      }
    }
  }

  foreach($ScanId in $ScanIdsReady)
  {
    $Uri =  "https://api.powerbi.com/v1.0/myorg/admin/workspaces/scanResult/$ScanId"
    $Response = Invoke-PowerBIRestMethod -Url $Uri -Method Get | ConvertFrom-Json
    $Response | ConvertTo-Json -depth 100 | Out-File "$OutputDirectory\WorkspaceMetadataScanStatus.json" -Append
  }
}


# Workspace Users
# https://learn.microsoft.com/en-us/power-bi/collaborate-share/service-roles-new-workspaces
# Do not use [Microsoft.PowerBI.Common.Api.Workspaces.Workspace[]] as type for parameter. Use generic PSObject. It is more forgiving when the API changes.
# MS had a bug in their code for hasWorkspaceLevelSettings field. Caused this error: Cannot process argument transformation on parameter 'Workspaces'. Cannot convert value...
function Export-WorkspacesUsers {
  param(
     $Workspaces
  )
  Write-Host "ExportWorkspacesUsers"
  foreach ($Workspace in $Workspaces) {
    $WorkspacesUsersReg = $Workspace | Select-Object -Property @{ name = "Workspace Id"; expression = { ($Workspace).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } } -ExpandProperty Users
    $WorkspacesUsersReg | Export-Csv "$OutputDirectory\WorkspacesUsers.csv" -Append
  }
}


function Export-WorkspacesDatasets {
  param(
    $Workspaces
  )

  Write-Host "ExportWorkspacesDatasets"

  foreach ($Workspace in $Workspaces) {
    $WorkspacesUsersReg = $Workspace | Select-Object -Property @{ name = "Workspace Id"; expression = { ($Workspace).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } } -ExpandProperty Datasets
    $WorkspacesUsersReg | Export-Csv "$OutputDirectory\WorkspacesDatasets.csv" -Append
  }
}



function Export-WorkspacesReports {
  param(
    $Workspaces
  )
  Write-Host "ExportWorkspacesReports"
  foreach ($Workspace in $Workspaces) {
    $WorkspacesUsersReg = $Workspace | Select-Object -Property @{ name = "Workspace Id"; expression = { ($Workspace).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } } -ExpandProperty Reports
    $WorkspacesUsersReg | Export-Csv "$OutputDirectory\WorkspacesReports.csv" -Append
  }
}


function Export-WorkspacesWorkbooks {
  param(
    $Workspaces
  )
  Write-Host "ExportWorkspacesWorkbooks"
  foreach ($Workspace in $Workspaces) {
    $WorkspacesUsersReg = $Workspace | Select-Object -Property @{ name = "Workspace Id"; expression = { ($Workspace).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } } -ExpandProperty Workbooks
    $WorkspacesUsersReg | Export-Csv "$OutputDirectory\WorkspacesWorkbooks.csv" -Append
  }
}


# This is call only returns the worksaces visible in app.powerbi.com
# Returns a list of workspaces the user has access to.
# GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/users
function Export-WorkspacesUsersREST {
  param(
    $Workspaces
  )

  $NonPersonalWorkspaces = $Workspaces | Where-Object type -CContains "Workspace"

  foreach ($Workspace in $NonPersonalWorkspaces) {
    $Uri = "https://api.powerbi.com/v1.0/myorg/groups/" + ($Workspace).ID.ToString() + "/users"
    Write-Host $Uri
    $Response = Invoke-PowerBIRestMethod -Url $Uri -Body $Body -Method Get | ConvertFrom-Json
    
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property @{ name = "Workspace Id"; expression = { ($Workspace).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\WorkspacesUsers-REST.csv" -Append
    }
  }

}


# Get the Reports and set the global variable.
$Reports = Get-PowerBIReport -Scope Organization
$Reports | Export-Csv "$OutputDirectory\Reports.csv"

$UniqueReportUsers = [System.Collections.Generic.List[reportUser]]::new() 

# Returns a list of users that have access to the specified report.
# Do not use strong typed parameters. The API changes and does not match the typings. :-(
# GET https://api.powerbi.com/v1.0/myorg/admin/reports/{reportId}/users

function Export-ReportUsersREST {
  param(
    $Reports
  )

  Write-Host "ExportReportUsersREST"
  #$UserIds = [System.Collections.Generic.List[String]]::new()

  foreach ($Report in $Reports) {
    $Uri = "https://api.powerbi.com/v1.0/myorg/admin/reports/" + $Report.Id.ToString() + "/users/"
    Write-Host $Uri
    $Response = Invoke-PowerBIRestMethod -Url $Uri -Body $Body -Method Get | ConvertFrom-Json
    
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $ReportUsers = $Response.value
      
      foreach ($ReportUser in $ReportUsers) {
        #deduplicate 
        #$result = $UserIds -contains $ReportUser.identifier
        $SearchResult = $UniqueReportUsers | Where-Object identifier -CContains $ReportUser.identifier

        if ($null -eq $SearchResult){
          $UniqueReportUsers.Add([reportUser]$ReportUser)
        }

        [reportUser]$ReportUser | Select-Object -Property @{ name = "Report Id"; expression = { ($Report).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\ReportsUsers-REST.csv" -Append 
      } 
    }
  }

}


# GET https://api.powerbi.com/v1.0/myorg/admin/users/{userId}/subscriptions
# Requires a call to Export-ReportUsersREST
function Export-ReportSubscriptionsUsersREST {
  param(
    [reportUser[]] $ReportUsers
  )
  Write-Host "Calling ExportReportSubscriptionsUsersREST"

  $stub = [ReportSubscriptionsUser]::new()
  $stubArray = [System.Collections.ArrayList]::new()
  $stubArray.Add($stub)
  $stubArray | Export-Csv "$OutputDirectory\ReportSubscriptionsUsers-REST.csv"

  # Filter out the system accounts.
  foreach ($ReportUser in $ReportUsers | Where-Object userType -eq "Member") {
    $url = "https://api.powerbi.com/v1.0/myorg/admin/users/" + $ReportUser.graphId.ToString() + "/subscriptions"
    $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
    
    if ($null -ne $Response.subscriptionEntities -and $Response.subscriptionEntities -gt 0) {
      $ResponseEntities = $Response.subscriptionEntities

      foreach ($ResponseEntity in $ResponseEntities) {
        $ResponseEntity | Select-Object -Property id,	title, artifactId,	artifactDisplayName, subArtifactDisplayName, artifactType, isEnabled,	frequency, startDate,	endDate, linkToContent,	previewImage,	attachmentFormat,	owner, @{ name = 'RecipientUsersExpanded'; expression = { $_.users -join ', ' }} | Export-Csv "$OutputDirectory\ReportSubscriptionsUsers-REST.csv" -Append 
      }
    }
  }

  $ContinuationUri = $Response.continuationUri

  while ($null -ne $ContinuationUri) {
    $Response = Invoke-PowerBIRestMethod -Url $ContinuationUri -Method Get | ConvertFrom-Json
    
    if ($null -ne $Response.subscriptionEntities -and $Response.subscriptionEntities -gt 0) {
      $ResponseEntities = $Response.subscriptionEntities

      foreach ($ResponseEntity in $ResponseEntities) {
        $ResponseEntity | Select-Object -Property id,	title, artifactId,	artifactDisplayName, subArtifactDisplayName, artifactType, isEnabled,	frequency, startDate,	endDate, linkToContent,	previewImage,	attachmentFormat,	owner, @{ name = 'RecipientUsersExpanded'; expression = { $_.Users -join ', ' } } | Export-Csv "$OutputDirectory\ReportSubscriptionsUsers-REST.csv" -Append 
      }
    }

    $ContinuationUri = $Response.continuationUri
  }
}

# GET GET https://api.powerbi.com/v1.0/myorg/groups/{groupId}/reports
# Reports in a workspace

function Export-WorkspaceReportsREST {
  param(
    $Workspaces
  )
  
  Write-Host "ExportWorkspaceReportsREST"
  $NonPersonalWorkspaces = $Workspaces | Where-Object type -Match "Workspace"

  foreach ($Workspace in $NonPersonalWorkspaces) {
    $Uri = "https://api.powerbi.com/v1.0/myorg/groups/" + ($Workspace).ID.ToString() + "/reports"
    Write-Host "Workspace: " ($Workspace).Name  
    $Response = Invoke-PowerBIRestMethod -Url $Uri -Body $Body -Method Get | ConvertFrom-Json

    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property id,	reportType,	name,	webUrl,	embedUrl,	isFromPbix,	isOwnedByMe,	datasetId,	datasetWorkspaceId, users, subscriptions | Export-Csv "$OutputDirectory\WorkspaceReports-REST.csv" -Append
    }
  }

}

# YOU CANNOT USE Select-Object -Property *.  When new fields are added, it breaks the CSV/Power Query import.

# TODO: Redo this call using admin rights. The PWSH cmd does not return all of the capcacities.
# https://api.powerbi.com/v1.0/myorg/admin/capacities
$Capacity = Get-PowerBICapacity 
$Capacity |  Select-Object -Property *, @{ name = 'AdminExpanded'; expression = { $_.Admins -join ' ' } } | Export-Csv "$OutputDirectory\Capacity.csv"

# GET https://api.powerbi.com/v1.0/myorg/capacities

function Export-PowerBICapacityREST {

  Write-Host "Calling Export-PowerBICapacityREST"
  
    $url = "https://api.powerbi.com/v1.0/myorg/capacities"
    $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Export-Csv "$OutputDirectory\Capacity-REST.csv" -Append
    }
  
}

function Export-PowerBICapacityAsAdminREST {

  Write-Host "Calling Export-PowerBICapacityAsAdminREST"
  
    $url = "https://api.powerbi.com/v1.0/myorg/admin/capacities?$expand=tenantKey"
    $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property id, displayName, sku, state, capacityUserAccessRight, region, tenantKey, @{ name = 'admins'; expression = { $_.admins -join ' ' } }, @{ name = 'users'; expression = { $_.users -join ' ' } } | Export-Csv "$OutputDirectory\CapacityAdmin-REST.csv" -Append
    }
  
}


$Dashboards = Get-PowerBIDashboard 
$Dashboards | Export-Csv "$OutputDirectory\Dashboard.csv"

function Export-DashboardTiles {
  param(
    $Dashboards
  )
  
  Write-Host "ExportDashboardTiles"
  foreach ($Dashboard in $Dashboards) {
    Get-PowerBITile -DashboardId $Dashboard.Id.ToString() | Export-Csv "$OutputDirectory\DashboardTiles.csv" -Append
  }

}

#GET https://api.powerbi.com/v1.0/myorg/reports/{reportId}/pages

function Export-ReportPagesREST {
  param(
    [Microsoft.PowerBI.Common.Api.Reports.Report[]] $Reports
  )
  Write-Host "ExportReportPagesREST"

  foreach ($Report in $Reports) {
    $Uri = "https://api.powerbi.com/v1.0/myorg/reports/" + ($Report).ID.ToString() + "/pages"
    Write-Host $Uri
    $Response = Invoke-PowerBIRestMethod -Url $Uri -Body $Body -Method Get | ConvertFrom-Json

    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property @{ name = "reportId"; expression = { ($Report).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\ReportPages-REST.csv" -Append
    }
  }

}


# Activity Events
# You must use the same date.
# $DayParameter should be multiples of 7 up, but a max of 30 days.
[Int32] $DayParameter = [System.Math]::Abs(14)

# This API has limited data returned. The REST version returns more data.
function Export-PowerBIActivityEvents {
  Write-Host "Calling ExportPowerBIActivityEvents"
  [string] $StartDateTime
  [string] $EndDateTime

  [datetime] $today = Get-Date
  [datetime] $StartDateTimeDT = $today.AddDays($DayParameter * -1)
  [datetime] $EndDateTimeDT = $StartDateTimeDT

  # Need to make sure not include 8 days. Otherwise, otherwise you double count the start day of the week. 
  $NumberOfDaysMaxIterator = $DayParameter

  for ($x = 1; $x -le $NumberOfDaysMaxIterator; $x++) {
    $StartDateTime = $StartDateTimeDT.AddDays($x).ToString("yyyy-MM-dd") + "T00:00:00"
    $EndDateTime = $EndDateTimeDT.AddDays($x).ToString("yyyy-MM-dd") + "T23:59:00"
    $Response = Get-PowerBIActivityEvent -StartDateTime $StartDateTime -EndDateTime $EndDateTime | ConvertFrom-Json
    $Response | Export-Csv "$OutputDirectory\ActivityEvents.csv" -Append -Force
  }
}

# GET https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime={startDateTime}&endDateTime={endDateTime}&continuationToken={continuationToken}&$filter={$filter}
# https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='2019-08-13T07:55:00.000Z'&endDateTime='2019-08-13T08:55:00.000Z'
# The Power BI events listed in Search the audit log in the Office 365 Protection Center will use this schema.
# https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-schema#power-bi-schema
# Definiton of the activity event.
# https://learn.microsoft.com/en-us/power-bi/enterprise/service-admin-auditing
function Export-PowerBIActivityEventsREST {
  Write-Host "Calling Export-PowerBIActivityEventsREST"
  [string] $StartDateTime
  [string] $EndDateTime
  [datetime] $today = Get-Date

  # Set the start back using DayParameter
  [datetime] $StartDateTimeDT = $today.AddDays($DayParameter * -1)
  [datetime] $EndDateTimeDT = $StartDateTimeDT
  $NumberOfDaysMaxIterator = $DayParameter

  # Actiivty Events responses objects come in different shapes and number of properties. I had to create 
  # super set of properties to avoid serialization errors. Utilitizng the folder import feature in Power Query. It uses sampling to build the import model. 
  # Create a serialized version of the object that can be filtered out in Power Query as the first row.
  # Populate all of the object properties with a ANY value so Power Query will create a column for it initially. Avoids ingest errors.

  $ActivityEventStub = [ActivityEventEntity]::new()
  $ActivityEventStub | Get-Member -MemberType Property | Where-Object Definition -Like "string *" | ForEach-Object {$ActivityEventStub.GetType().GetProperty($_.Name).SetValue($ActivityEventStub, "0_DELETETHIS")}
  $ActivityEventStub | Get-Member -MemberType Property | Where-Object Definition -Like "datetime*" | ForEach-Object {$ActivityEventStub.GetType().GetProperty($_.Name).SetValue($ActivityEventStub, $today)}
  $ActivityEventStub | Get-Member -MemberType Property | Where-Object Definition -Like "System.Object*" | ForEach-Object {$ActivityEventStub.GetType().GetProperty($_.Name).SetValue($ActivityEventStub, @())}
  $ActivityEventStub.ModelsSnapshots = @()

  CreateDirectoryIfNotExists("$OutputDirectory\ActivityEvents")

  for ($x = 1; $x -le $NumberOfDaysMaxIterator; $x++) {
    $StartDateTime = $StartDateTimeDT.AddDays($x).ToString("yyyy-MM-dd") + "T00:00:00.000Z"
    $EndDateTime = $EndDateTimeDT.AddDays($x).ToString("yyyy-MM-dd") + "T23:59:00.000Z"

    $url = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='$StartDateTime'&endDateTime='$EndDateTime'"
    $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
    
    if ($null -ne $Response.activityEventEntities -and $Response.activityEventEntities.Length -gt 0) {
      # Add the stub to the initial set of files.
      # Ensures the output will be a list and not a record. 
      # Avoids the Power Query error. 'Expression.Error: We cannot convert a value of type Record to type List'
      $ActivityEventResponses = [System.Collections.ArrayList]::new()
      #Create multiple stubs in case there is no data on that day.
      $ActivityEventResponses.Add($ActivityEventStub)
      $ActivityEventResponses.Add($ActivityEventStub)

      $ActivityEventResponses.AddRange($Response.activityEventEntities)
      $CleanName = (Get-Date).ToString('yyyy-MM-ddThh-mm-ss-fff')
      
      # Need to create a JSON file because the PowerShell auto-mapping conversion to CSV will drop columns it fails to convert. 
      $ActivityEventResponses | ConvertTo-Json -depth 100 | Out-File "$OutputDirectory\ActivityEvents\$CleanName-ActivityEvents-REST.json" -Append
    }

    $ContinuationUri = $Response.continuationUri

    while ($null -ne $ContinuationUri) {
      $Response = Invoke-PowerBIRestMethod -Url $ContinuationUri -Method Get | ConvertFrom-Json

      if ($null -ne $Response.activityEventEntities -and $Response.activityEventEntities.Length -gt 0) {
        $ActivityEventResponses = [System.Collections.ArrayList]::new()
        $ActivityEventResponses.Add($ActivityEventStub)
        $ActivityEventResponses.Add($ActivityEventStub)
        $ActivityEventResponses.AddRange($Response.activityEventEntities)
        $CleanName = (Get-Date).ToString('yyyy-MM-ddThh-mm-ss-fff')
        $ActivityEventResponses | ConvertTo-Json -depth 100 | Out-File "$OutputDirectory\ActivityEvents\$CleanName-ActivityEvents-REST.json" -Append
      }
  
      $ContinuationUri = $Response.continuationUri
    }

  }
}

# https://learn.microsoft.com/en-us/rest/api/power-bi/admin/widely-shared-artifacts-links-shared-to-whole-organization
# GET https://api.powerbi.com/v1.0/myorg/admin/widelySharedArtifacts/linksSharedToWholeOrganization

function Export-LinksSharedToWholeOrganizationREST {
  Write-Host "Calling ExportLinksSharedToWholeOrganizationREST"

  $url = "https://api.powerbi.com/v1.0/myorg/admin/widelySharedArtifacts/linksSharedToWholeOrganization"
  $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
  #@odata.context	days	times	enabled	localTimeZoneId	notifyOption
  if ($null -ne $Response.artifactAccessEntities) {
    $ArtifactAccessEntities = $Response.artifactAccessEntities

    foreach ($ArtifactAccessEntity in $ArtifactAccessEntities) {
      [SharedLink]$ArtifactAccessEntity | Select-Object -Property artifactId,	artifactType,	accessRight,	shareType, @{name = "ResourceDisplayName"; expression = { $_.displayName } } -ExpandProperty sharer | Export-Csv "$OutputDirectory\SharedLinks-REST.csv" -Append 
    }
  }

  $ContinuationUri = $Response.continuationUri

  while ($null -ne $ContinuationUri) {
    $Response = Invoke-PowerBIRestMethod -Url $ContinuationUri -Method Get | ConvertFrom-Json
    #@odata.context	daystimes	enabled	localTimeZoneId	notifyOption
    if ($null -ne $Response.artifactAccessEntities) {
      foreach ($ArtifactAccessEntity in $ArtifactAccessEntities) {
        [SharedLink]$ArtifactAccessEntity | Select-Object -Property artifactId,	artifactType,	accessRight,	shareType, @{name = "ResourceDisplayName"; expression = { $_.displayName } } -ExpandProperty sharer | Export-Csv "$OutputDirectory\SharedLinks-REST.csv" -Append 
      }
    }

    $ContinuationUri = $Response.continuationUri
  }
}


#GET https://api.powerbi.com/v1.0/myorg/admin/widelySharedArtifacts/publishedToWeb

function Export-WidelySharedArtifactsPublishedToWebREST {
  Write-Host "Calling ExportLinksSharedToWholeOrganizationREST"
  # Create a blank row to show there were no shared web links.
  New-Object SharedLink | Select-Object -Property artifactId,	artifactType,	accessRight,	shareType, @{name = "ResourceDisplayName"; expression = { $_.displayName } } -ExpandProperty sharer | Export-Csv "$OutputDirectory\PublishedToWeb-REST.csv"

  $url = "https://api.powerbi.com/v1.0/myorg/admin/widelySharedArtifacts/publishedToWeb"
  $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
  #@odata.context	days	times	enabled	localTimeZoneId	notifyOption
  if ($null -ne $Response.artifactAccessEntities) {
    $ArtifactAccessEntities = $Response.artifactAccessEntities

    foreach ($ArtifactAccessEntity in $ArtifactAccessEntities) {
      [SharedLink]$ArtifactAccessEntity | Select-Object -Property artifactId,	artifactType,	accessRight,	shareType, @{name = "ResourceDisplayName"; expression = { $_.displayName } } -ExpandProperty sharer | Export-Csv "$OutputDirectory\PublishedToWeb-REST.csv" -Append 
    }
  }

  $ContinuationUri = $Response.continuationUri

  while ($null -ne $ContinuationUri) {
    $Response = Invoke-PowerBIRestMethod -Url $ContinuationUri -Method Get | ConvertFrom-Json
    #@odata.context	days	times	enabled	localTimeZoneId	notifyOption
    if ($null -ne $Response.artifactAccessEntities) {
      foreach ($ArtifactAccessEntity in $ArtifactAccessEntities) {
        [SharedLink]$ArtifactAccessEntity | Select-Object -Property artifactId,	artifactType,	accessRight,	shareType, @{name = "ResourceDisplayName"; expression = { $_.displayName } } -ExpandProperty sharer | Export-Csv "$OutputDirectory\PublishedToWeb-REST.csv" -Append 
      }
    }

    $ContinuationUri = $Response.continuationUri
  }
}

# GET https://api.powerbi.com/v1.0/myorg/admin/dataflows
function Export-DataflowAsAdmin {

  Write-Host "Calling Export-DataflowAsAdmin"
    $url = "https://api.powerbi.com/v1.0/myorg/admin/dataflows"
    $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
    $Response.value | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Dataflows-REST.csv"

}


# Returns a list of Power BI items (such as reports or dashboards) that the specified user has access to.
# GET https://api.powerbi.com/v1.0/myorg/admin/users/{userId}/artifactAccess
# This call does not work if your data Workspace object version does not match this call's version. 
# Dataflows
function Export-PowerBIDataflows {
  param(
    [Microsoft.PowerBI.Common.Api.Workspaces.Workspace[]]$Workspaces
  )

  Write-Host "Calling ExportPowerBIDataflows"
  foreach ($Workspace in $Workspaces) {
    # Converts to the Microsoft.PowerBI.Common.Api.Workspaces.Workspace type.
    Get-PowerBIDataFlow -Workspace $Workspace | ConvertTo-Csv | Out-File -FilePath "$OutputDirectory\DataFlow.csv" -Append
  }
}

# This call does not return all of the datasets. It only returns the datasets your user account has been assigned to.  To get all of the datasets, you need to call AsAdmin API.
#  $Datasets2 = Get-PowerBIDataset -Scope Organization -Include actualStorage
#  $Datasets2 | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Datasets3.csv"

$Datasets

# Use this call because the admin call brings back all tenant data.
# GET https://api.powerbi.com/v1.0/myorg/admin/datasets
function Export-PowerBIDatasetAsAdminREST {
  Write-Host "Calling ExportPowerBIDatasetAsAdminREST"
  $url = "https://api.powerbi.com/v1.0/myorg/admin/datasets"
  $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
  $global:Datasets = $Response.value
  $global:Datasets | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Datasets.csv"
}


# This call does not work.  It does not return any records. DNU.
function Export-Dataset2REST {

  Write-Host "Calling ExportDataset2REST"
    $url = "https://api.powerbi.com/v1.0/myorg/datasets"
    $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
    $Response.value | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Dataset2-REST.csv"
}

# Dataset Sources
# Analysis Services Backup (ABF)

function Export-PowerBIDatasources {
  param(
     $Datasets
  )
  Write-Host "Calling ExportPowerBIDatasources"
  foreach ($Dataset in $Datasets) {
    $Uri = ($Dataset).Id.ToString()
    $Datasource = Get-PowerBIDatasource -DatasetId $Uri -Scope Organization | Select-Object -Property @{ name = "datasetId"; expression = { ($Dataset).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } } , * -ExpandProperty ConnectionDetails
    $Datasource | Export-Csv "$OutputDirectory\Datasources.csv" -Append
  }
}


function Export-PowerBIDatasourcesREST {
  param(
     $Datasets
  )

  Write-Host "Calling ExportPowerBIDatasourcesREST"
  
  foreach ($Dataset in $Datasets) {
    Write-Host ($Dataset).Id.ToString()
    $url = "https://api.powerbi.com/v1.0/myorg/datasets/" + ($Dataset).Id.ToString() + "/datasources"
    $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property @{Name = "Dataset ID"; Expression = { $Dataset.Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Datasources-REST.csv" -Append
    }
  }
}

# DO NOT USE.
function Export-PowerBIDatasetStorage {
  param(
     $Datasets
  )
  return
  Write-Host "Calling ExportPowerBIDatasetStorage"
  foreach ($Dataset in $Datasets) {
    $Dataset | Select-Object -Property Id -ExpandProperty ActualStorage | Export-Csv "$OutputDirectory\DatasetStorage.csv" -Append
  }
}

# GET https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/refreshSchedule
# Errors "Dataset e6e823bc-7384-4cd3-a972-22a275a14539 is not found! please verify datasetId is correct and user have sufficient permissions." 
function Export-DatasetRefreshSheduleREST {
  param(
     $Datasets
  )
  
  Write-Host "Calling ExportDatasetRefreshSheduleREST"

  foreach ($Dataset in $Datasets) {
    #Write-Host ($Dataset).Id.ToString()
    $url = "https://api.powerbi.com/v1.0/myorg/datasets/" + ($Dataset).Id.ToString() + "/refreshSchedule"
    $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json
    #@odata.context	days	times	enabled	localTimeZoneId	notifyOption
    if ($null -ne $Response) {
      $Response | Select-Object -Property @{ name = "datasetId"; expression = { ($Dataset).Id.ToString() } }, @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, localTimeZoneId, @{name = "days"; expression = { $_.days } }, @{name = "times"; expression = { $_.times } }, enabled, notifyOption | Export-Csv "$OutputDirectory\DatasetRefreshSchedule-REST.csv" -Append
    }
  }
}


# GET https://api.powerbi.com/v1.0/myorg/datasets/{datasetId}/refreshes
# The API can only be called on a Model-based dataset, which is a dataset based on an imported data model. DirectQuery and 
# Live Connection datasets do not have a data model, so you cannot retrieve refresh history for them using this API.
# TODO: Move towards using GET https://api.powerbi.com/v1.0/myorg/admin/capacities/refreshables

function Export-DatasetRefreshHistoryREST {
  param(
    $Datasets
  )

  Write-Host "Calling ExportDatasetRefreshHistoryREST"

  foreach ($Dataset in $Datasets) {
    Write-Host $Dataset.Name " DatasetId: " $Dataset.Id
    
    if ($null -ne $Dataset.Id) {
      $url = "https://api.powerbi.com/v1.0/myorg/datasets/" + ($Dataset).Id.ToString() + "/refreshes"
      $Response = Invoke-PowerBIRestMethod -Url $url -Method Get | ConvertFrom-Json

      if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
        $datasetRefreshHistoryItem = [DatasetRefreshHistoryItem[]]$Response.value
        $datasetRefreshHistoryItem | Select-Object -Property @{ name = "Dataset Id"; expression = { ($Dataset).Id.ToString() } }, * | Export-Csv "$OutputDirectory\DatasetRefreshHistoryREST.csv" -Append
      }
    }
  }
}

# Apps
#GET https://api.powerbi.com/v1.0/myorg/apps
# Need to create file with the basic fields
$Apps = [System.Collections.ArrayList]::new()

# Set the $Apps variable so it can be used for later calls.
function Export-PowerBIAppsREST {
  Write-Host "Calling ExportPowerBIAppsREST"
  
  # Create a file with empty rows
  $AppStub = [App]::new()
  $global:Apps.Add($AppStub)
  $global:Apps | Export-Csv "$OutputDirectory\Apps-REST.csv"

  # Top call is requied by the API. Set the number to very large number.
  $url = 'https://api.powerbi.com/v1.0/myorg/admin/apps?$top=200'
  $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
  if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
    $global:Apps.AddRange($Response.value)
    $global:Apps | Export-Csv "$OutputDirectory\Apps-REST.csv" -Append
  } 
}


#GET https://api.powerbi.com/v1.0/myorg/admin/apps/{appId}/users

function Export-AppUsersREST {
  param($Apps)
  
  Write-Host "ExportAppUsersREST"

  $AppUserStub = [AppUser]::new()
  $AppUsers = [System.Collections.ArrayList]::new()
  $AppUsers.Add($AppUserStub)
  $AppUsers | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, @{ name = "App Id"; expression = {($App.id).ToString()}}, * | Export-Csv "$OutputDirectory\AppsUsers-REST.csv"

  foreach ($App in $Apps) {
    $url = "https://api.powerbi.com/v1.0/myorg/admin/apps/" + $App.id + "/users"
    $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
    if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
      $Response.value | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, @{ name = "App Id"; expression = {($App.id).ToString()}}, * | Export-Csv "$OutputDirectory\AppsUsers-REST.csv" -Append 
    }

  }

}

#GET https://api.powerbi.com/v1.0/myorg/gateways

function Export-GatewaysREST {

  Write-Host "Calling ExportGatewaysREST"

  $url = "https://api.powerbi.com/v1.0/myorg/gateways"
  $Response = Invoke-PowerBIRestMethod -Url $url -Body $Body -Method Get | ConvertFrom-Json
  
  if ($null -ne $Response.value -and $Response.value.Length -gt 0) {
    $Gateways = $Response.value
    $Gateways | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\Gateways-REST.csv" 
  }

}


# Get-PowerBITable -DatasetId eed49d27-8e3c-424d-9342-c6b3ca6db64d
# Only works for Push Datasets - Datasets PostDatasetInGroup. Not very useful.
# Get-PowerBIDataset | ? AddRowsApiEnabled -eq $true | Get-PowerBITable
function Export-DatasetTables {
  param(
     $Datasets
  )

  Write-Host "Calling ExportDatasetTables"

  foreach ($Dataset in $Datasets) {
    Get-PowerBITable -DatasetId ($Dataset).Id.ToString() | Select-Object -Property @{Name = "Date Retrieved"; Expression = { $RetrieveDate } }, * | Export-Csv "$OutputDirectory\DatasetTables.csv" -Append
  }
}

# POST https://api.powerbi.com/v1.0/myorg/admin/workspaces/getInfo?lineage={lineage}&datasourceDetails={datasourceDetails}&datasetSchema={datasetSchema}&datasetExpressions={datasetExpressions}&getArtifactUsers={getArtifactUsers}


# TODO: Create explicit iterfaces for response exports and cast to them. When Microsoft updates the REST API and the amount of properties vary in the REST response, it
# can cause the Export-Csv to fail depending if the first item has less properties than the subsequent response items. Avoid SelectObject *. It may fail later.
# Error description "The appended object does not have a property that corresponds to the following column ... "

class ActivityEventEntity {
  [string] $Activity
  [string] $ActivityId
  [System.Management.Automation.PSCustomObject] $AggregatedWorkspaceInformation
  [string] $AppName
  [string] $AppReportId
  [string] $ArtifactId
  [string] $ArtifactKind
  [string] $ArtifactName
  [string] $ArtifactObjectId
  [string] $CapacityId
  [string] $CapacityName
  [string] $CapacityState
  [string] $CapacityUsers
  [string] $ClientIP
  [string] $ConsumptionMethod
  [datetime] $CreationTime
  [string] $DataConnectivityMode
  [string] $DatasetId
  [string] $DatasetName
  [string] $DistributionMethod
  [datetime] $ExportEventEndDateTimeParameter
  [datetime] $ExportEventStartDateTimeParameter
  [string] $FolderDisplayName
  [string] $FolderObjectId
  [string] $Id
  [string] $ImportDisplayName
  [string] $ImportId
  [string] $ImportSource
  [string] $ImportType
  [bool] $IsSuccess
  [string] $ItemName
  [datetime] $LastRefreshTime
  [PSObject[]]$ModelsSnapshots
  [string] $ObjectId
  [string] $Operation
  [string] $OrganizationId
  [string]$PaginatedReportDataSources
  [Int32] $RecordType
  [Int32] $RefreshEnforcementPolicy
  [string] $RefreshType
  [string] $ReportId
  [string] $ReportName
  [string] $ReportType
  [string] $RequestId
  [string] $UserAgent
  [string] $UserId
  [string] $UserKey
  [Int32] $UserType
  [string] $Workload
  [string] $WorkspaceId
  [string] $WorkSpaceName
  [string] $WorkspacesSemicolonDelimitedList
}


class App {
  $id
  $name
  $lastUpdate
  $description
  $publishedBy
  $workspaceId
  $users
}


class AppUser {
  $appUserAccessRight
  $emailAddress
  $displayName
  $identifier
  $graphId
  $principalType
  $userType
}


class SharedLink {
  $artifactId	
  $displayName	
  $artifactType	
  $accessRight	
  $shareType	
  [sharer]$sharer
}

class Sharer {
  $displayName
  $emailAddress
  $identifier
  $graphId
  $principalType
}


class ReportUser {
  $displayName
  $emailAddress	
  $graphId	
  [string] $identifier	
  $principalType	
  $profile	
  $reportUserAccessRight	
  $userType	
}

class ReportSubscriptionsUser {
  $id
  $title
  $artifactId
  $artifactDisplayName
  $subArtifactDisplayName
  $artifactType
  $isEnabled
  $frequency
  $startDate
  $endDate
  $linkToContent
  $previewImage
  $attachmentFormat
  $owner
  $RecipientUsersExpanded
}



class DatasetRefreshHistoryItem {
  $endTime
  $id
  $refreshAttempts
  $refreshType
  $requestId
  $serviceExceptionJson
  $startTime
  $status
}


# These functions must always be called.
Export-WorkspacesAsAdmin
Export-PowerBIDatasetAsAdminREST
Export-ReportUsersREST($Reports)

# Optionally called.

Export-PowerBICapacityREST
Export-PowerBICapacityAsAdminREST
Export-AppUsers($Apps)
Export-DashboardTiles($Dashboards)
Export-DatasetRefreshHistoryREST($Datasets)
Export-DatasetRefreshSheduleREST($Datasets)
Export-DataflowAsAdmin
Export-LinksSharedToWholeOrganizationREST
Export-PowerBIActivityEventsREST
Export-PowerBIAppsREST
Export-AppUsersREST($Apps)
Export-PowerBIDatasetStorage($Datasets)
Export-PowerBIDatasources($Datasets)
Export-PowerBIDatasourcesREST($Datasets)
Export-ReportPagesREST($Reports)
Export-ReportSubscriptionsUsersREST($UniqueReportUsers)
Export-WorkspaceReportsREST($Workspaces)
Export-WorkspacesDatasets($Workspaces)
Export-WorkspacesReports($Workspaces)
Export-WorkspacesUsers($Workspaces)
Export-WorkspacesUsersREST($Workspaces)
Export-WorkspacesWorkbooks($Workspaces)

# DO NOT USE
# Export-PowerBIActivityEvents
# Export-Dataset2REST - Do not use.
# Export-WidelySharedArtifactsPublishedToWebREST
# PostWorkspaceInfo($Workspaces)
# Export-DatasetTables($Datasets)
# Export-GatewaysREST
# Export-PowerBIDataflows
Stop-Transcript