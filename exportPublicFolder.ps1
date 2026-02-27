<#
.SYNOPSIS
  # Ensure destination folder exists and match folder type (calendar/tasks/etc.)
  try {
    $dest = $DestParentFolder.Folders.Item($SourceFolder.Name)
  } catch {
    $dest = New-DestFolderMatchingType -DestParentFolder $DestParentFolder -SourceFolder $SourceFolder -WhatIf:$WhatIf
  }
.PARAMETER EwsDllPath
  Path to Microsoft.Exchange.WebServices.dll (EWS Managed API 2.2)

.PARAMETER EwsUrl
  Optional override for EWS endpoint (e.g. https://outlook-dod.office365.us/EWS/Exchange.asmx)

.PARAMETER PstFileName
  PST file name to create on Desktop.

.PARAMETER WhatIf
  If set, no copy actions occur (enumeration only).

.NOTES
  - EWS enumerates folders; Outlook MAPI/COM creates PST and copies items/folders.
  - Requires Outlook profile with access to Public Folders.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)]
  [string]$PublicFolderPath,

  [string]$EwsDllPath = "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Exchange.WebServices.2.2\lib\40\Microsoft.Exchange.WebServices.dll",

  [string]$EwsUrl,

  [string]$PstFileName = "PublicFolderExport.pst",

  [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
  param([string]$Message, [string]$Level = "INFO")
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$ts][$Level] $Message"
}

function New-DestFolderMatchingType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $DestParentFolder,   # Outlook.Folder (PST)
        [Parameter(Mandatory)] $SourceFolder,       # Outlook.Folder (Public Folder)
        [switch] $WhatIf
    )

    # If it already exists, reuse it
    try {
        return $DestParentFolder.Folders.Item($SourceFolder.Name)
    } catch {}

    if ($WhatIf) {
        # In WhatIf mode, don't actually create; return parent as placeholder
        return $DestParentFolder
    }

    # Map SourceFolder.DefaultItemType (OlItemType) to Folder.Add Type (OlDefaultFolders)
    # DefaultItemType tells what item type the folder is meant to contain. [1](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/api/Outlook.OlExchangeStoreType.md)
    
    $addType = $null

    try {
        $srcItemType = [int]$SourceFolder.DefaultItemType
    } catch {
        $srcItemType = -1
    }

    switch ($srcItemType) {
        1 { $addType = 9  }   
        3 { $addType = 13 }   
        2 { $addType = 10 }   
        4 { $addType = 12 }   
        default { $addType = 6 } 
    }

    # Create folder of the correct type in the PST
    return $DestParentFolder.Folders.Add($SourceFolder.Name, $addType)
}


function Get-OutlookSession {
  # Uses the default Outlook profile (current user context).
  $outlook = New-Object -ComObject Outlook.Application
  $ns = $outlook.GetNamespace("MAPI")
  # If Outlook isn't running, this ensures session is initialized under default profile.
  $ns.Logon($null, $null, $false, $false) | Out-Null
  return $ns
}

function Get-PrimarySmtpFromOutlook {
  param([Parameter(Mandatory)]$Session)

  try {
    $ae = $Session.CurrentUser.AddressEntry
    $exUser = $ae.GetExchangeUser()
    if ($exUser -and $exUser.PrimarySmtpAddress) { return $exUser.PrimarySmtpAddress }
  } catch {}

  # Fallback: try current user address
  try {
    if ($Session.CurrentUser.Address) { return $Session.CurrentUser.Address }
  } catch {}

  throw "Unable to determine Primary SMTP from Outlook profile."
}

function Connect-Ews {
  param(
    [Parameter(Mandatory)][string]$SmtpAddress,
    [string]$UrlOverride
  )

  if (-not (Test-Path $EwsDllPath)) {
    throw "EWS DLL not found at: $EwsDllPath"
  }
  Add-Type -Path $EwsDllPath

  $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
  )

  # "Outlook profile" credential reuse is not directly exposed for EWS; 
  # this uses the current Windows identity (works for many on-prem/kerberos scenarios).
  $service.UseDefaultCredentials = $true

  if ($UrlOverride) {
    $service.Url = [Uri]$UrlOverride
    Write-Log "EWS URL override set to $UrlOverride"
  } else {
    # Autodiscover
    $service.AutodiscoverUrl($SmtpAddress, { param($redirectionUrl) return $redirectionUrl.ToLower().StartsWith("https://") })
    Write-Log "EWS autodiscover resolved to $($service.Url)"
  }

  return $service
}

function Resolve-EwsPublicFolderByPath {
  param(
    [Parameter(Mandatory)]$Service,
    [Parameter(Mandatory)][string]$Path
  )

  $clean = $Path.Trim()
  if ([string]::IsNullOrWhiteSpace($clean)) { throw "PublicFolderPath is empty." }
  if ($clean[0] -ne "\") { $clean = "\" + $clean }

  $segments = $clean.Trim("\").Split("\") | Where-Object { $_ -ne "" }

  $current = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
    $Service,
    [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::PublicFoldersRoot
  )

  foreach ($seg in $segments) {
    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
    $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
      [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $seg
    )
    $found = $Service.FindFolders($current.Id, $filter, $view).Folders | Select-Object -First 1
    if (-not $found) { throw "EWS: Folder segment not found: '$seg' under '$($current.DisplayName)'" }
    $current = $found
  }

  return $current
}

function Get-EwsFolderTree {
  param(
    [Parameter(Mandatory)]$Service,
    [Parameter(Mandatory)]$RootFolder,
    [Parameter(Mandatory)][string]$RootPath
  )

  $results = New-Object System.Collections.Generic.List[object]

    # Attempt to get Folder Item Count, but fall back to -1 if it fails (e.g. permissions issue, or certain folder types that don't support it).
  function Recurse($folder, $path) {
    try {
      $itemCount = [int]$folder.TotalCount
    } catch {
      $itemCount = -1
    }

    $results.Add([pscustomobject]@{
      Path        = $path
      DisplayName = $folder.DisplayName
      FolderId    = $folder.Id.UniqueId
      ItemCount   = $itemCount
    }) | Out-Null

    $view = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow

    $find = $Service.FindFolders($folder.Id, $view)
    foreach ($child in $find.Folders) {
      Recurse $child ("$path\$($child.DisplayName)")
    }
  }

  Recurse $RootFolder $RootPath
  return $results
}

function Get-PublicFolderStoreRoot {
  param([Parameter(Mandatory)]$Session)

  # olExchangePublicFolder is 2
  $publicStore = $null
  foreach ($store in $Session.Stores) {
    try {
      if ($store.ExchangeStoreType -eq 2) { $publicStore = $store; break }
    } catch {}
  }
  if (-not $publicStore) { throw "Outlook: Could not find an Exchange Public Folder store in the profile." }

  return $publicStore.GetRootFolder()
}

function Resolve-MapiFolderByPath {
  param(
    [Parameter(Mandatory)]$PublicRoot,
    [Parameter(Mandatory)][string]$Path
  )

  $clean = $Path.Trim()
  if ($clean[0] -ne "\") { $clean = "\" + $clean }
  $segments = $clean.Trim("\").Split("\") | Where-Object { $_ -ne "" }

  # Some orgs have "All Public Folders" as the first visible node.
  $current = $PublicRoot
  $allPf = $null
  try { $allPf = $PublicRoot.Folders.Item("All Public Folders") } catch {}

  if ($allPf) {
    # Prefer "All Public Folders" if the first segment isn't directly under root.
    $firstTry = $null
    try { $firstTry = $PublicRoot.Folders.Item($segments[0]) } catch {}
    if (-not $firstTry) { $current = $allPf }
  }

  foreach ($seg in $segments) {
    try {
      $current = $current.Folders.Item($seg)
    } catch {
      throw "Outlook: Folder segment not found: '$seg'"
    }
  }

  return $current
}

function Ensure-PstStoreAndRoot {
  param(
    [Parameter(Mandatory)]$Session,
    [Parameter(Mandatory)][string]$PstPath
  )

  # olStoreUnicode is 2 for AddStoreEx
  $Session.AddStoreEx($PstPath, 2) | Out-Null

  $pstStore = $null
  foreach ($store in $Session.Stores) {
    try {
      if ($store.FilePath -and ($store.FilePath -ieq $PstPath)) { $pstStore = $store; break }
    } catch {}
  }
  if (-not $pstStore) { throw "Outlook: PST store not found after AddStoreEx: $PstPath" }

  return $pstStore.GetRootFolder()
}

function Copy-ItemsToFolder {
  param(
    [Parameter(Mandatory)]$SourceFolder,
    [Parameter(Mandatory)]$DestFolder,
    [switch]$WhatIfMode
  )

  $items = $SourceFolder.Items
  $count = $items.Count
  Write-Log "Copying $count items from '$($SourceFolder.FolderPath)' -> '$($DestFolder.FolderPath)'"

  for ($i = 1; $i -le $count; $i++) {
    $item = $null
    try {
      $item = $items.Item($i)
      if (-not $item) { continue }

      if ($WhatIfMode) { continue }

      # Copy() returns a duplicate in the source folder, then Move() moves that duplicate to destination
      $copied = $item.Copy()
      $null = $copied.Move($DestFolder)

    } catch {
      Write-Log "Item copy failed at index $i in '$($SourceFolder.FolderPath)': $($_.Exception.Message)" "WARN"
    }
  }
}

# Global export stats (folder-level)
$script:ExportStats = New-Object System.Collections.Generic.List[object]

function Get-FolderItemCountSafe {
    param([Parameter(Mandatory)]$Folder)

    try {
        # Items.Count can be slow on very large folders but is the simplest and works for PF mail folders
        return [int]$Folder.Items.Count
    } catch {
        return -1
    }
}


function Copy-ItemsToFolderWithCount {
    param(
        [Parameter(Mandatory)]$SourceFolder,
        [Parameter(Mandatory)]$DestFolder,
	[switch] $WhatIf
    )

    $total     = 0
    $exported  = 0
    $failed    = 0

    try {
        $items = $SourceFolder.Items
        $total = $items.Count
    } catch {
        return @{
            Total    = -1
            Exported = 0
            Failed  = 0
        }
    }

    for ($i = 1; $i -le $total; $i++) {
        try {
            $item = $items.Item($i)
            if (-not $item) { continue }

            # Copy works for MailItem AND AppointmentItem
            $copied = $item.Copy()

            # Move the copied item into the destination folder. Move() returns the moved item
            # in the destination store; capture it so we can inspect/update properties.
            try {
              $moved = $copied.Move($DestFolder)
            } catch {
              $failed++
              continue
            }

            # After moving, Outlook may have altered the Subject (e.g. prefixed "Copy:").
            # If so, restore the original Subject on the moved item.
            try {
              if ($moved -and ($moved.Subject -ne $item.Subject)) {
                $moved.Subject = $item.Subject
                $moved.Save()
              }
            } catch {
              # Non-fatal; continue even if we cannot modify the moved item.
            }

            $exported++
        }
        catch {
            $failed++
        }
    }

    return @{
        Total    = $total
        Exported = $exported
        Failed  = $failed
    }
}

function Copy-FolderRecursiveWithStats {
    param(
        [Parameter(Mandatory)]$SourceFolder,
        [Parameter(Mandatory)]$DestParentFolder,
        [Parameter(Mandatory)][string]$SourceFolderPath,
	[switch] $WhatIf
    )

    # Ensure destination folder exists and match folder type (calendar/tasks/etc.)
    try {
      $dest = $DestParentFolder.Folders.Item($SourceFolder.Name)
    } catch {
      $dest = New-DestFolderMatchingType -DestParentFolder $DestParentFolder -SourceFolder $SourceFolder -WhatIf:$WhatIf
    }

    # Copy items (mail + calendar)
    $counts = Copy-ItemsToFolderWithCount `
        -SourceFolder $SourceFolder `
        -DestFolder   $dest -WhatIf:$WhatIf

    # Record stats for CSV
    $script:ExportStats.Add([pscustomobject]@{
        FolderPath        = $SourceFolderPath
        FolderName        = $SourceFolder.Name
        FolderId          = $SourceFolder.EntryID
        FolderType        = $SourceFolder.DefaultItemType
        SourceItemCount   = $counts.Total
        ExportedItemCount = $counts.Exported
        FailedItemCount   = $counts.Failed
    }) | Out-Null

    # Recurse into subfolders
    foreach ($sf in $SourceFolder.Folders) {
        $childPath = "$SourceFolderPath\$($sf.Name)"
        Copy-FolderRecursiveWithStats `
            -SourceFolder $sf `
            -DestParentFolder $dest `
            -SourceFolderPath $childPath -WhatIf:$WhatIf
    }
}

# Note: The original Copy-FolderRecursive is left here for reference; the new version with stats is above.

# function Copy-FolderRecursive {
#   param(
#     [Parameter(Mandatory)]$SourceFolder,
#     [Parameter(Mandatory)]$DestParentFolder,
#     [switch]$WhatIfMode
#   )

#   # Create destination folder (same display name)
#   $dest = $null
#   try {
#     $dest = $DestParentFolder.Folders.Item($SourceFolder.Name)
#   } catch {
#     if ($WhatIfMode) {
#       Write-Log "Would create folder '$($SourceFolder.Name)' under '$($DestParentFolder.FolderPath)'"
#       # Create a temporary placeholder object? We'll just skip item copies when WhatIf.
#       $dest = $DestParentFolder
#     } else {
#       $dest = $DestParentFolder.Folders.Add($SourceFolder.Name)
#       Write-Log "Created folder '$($SourceFolder.Name)' under '$($DestParentFolder.FolderPath)'"
#     }
#   }

#   if (-not $WhatIfMode) {
#     Copy-ItemsToFolder -SourceFolder $SourceFolder -DestFolder $dest -WhatIfMode:$false
#   }

#   # Recurse subfolders
#   foreach ($sf in $SourceFolder.Folders) {
#     Copy-FolderRecursive -SourceFolder $sf -DestParentFolder $dest -WhatIfMode:$WhatIfMode
#   }
# }


# =========================
# Main
# =========================

Write-Log "Starting export for PublicFolderPath: $PublicFolderPath"

# Outlook session from profile
$session = Get-OutlookSession
$smtp = Get-PrimarySmtpFromOutlook -Session $session
Write-Log "Using Outlook profile SMTP: $smtp"

# EWS connect + enumerate
$service = Connect-Ews -SmtpAddress $smtp -UrlOverride $EwsUrl
$ewsRoot = Resolve-EwsPublicFolderByPath -Service $service -Path $PublicFolderPath
Write-Log "EWS resolved folder: '$($ewsRoot.DisplayName)'"

$tree = Get-EwsFolderTree -Service $service -RootFolder $ewsRoot -RootPath $PublicFolderPath
$PFString = $PublicFolderPath.split("\")[1] # Get top-level PF name for log
$desktop = [Environment]::GetFolderPath("Desktop")
$listingPath = Join-Path $desktop ("Preview_exportPublicFolder_$PFString`_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
$tree | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $listingPath
Write-Log "Subfolder listing saved to: $listingPath"

# PST creation + copy via Outlook MAPI
$pstPath = Join-Path $desktop $PstFileName
Write-Log "PST target: $pstPath"

$publicRoot = Get-PublicFolderStoreRoot -Session $session
$srcFolder = Resolve-MapiFolderByPath -PublicRoot $publicRoot -Path $PublicFolderPath

# If WhatIf is enabled, we will still enumerate and create folders but skip actual item copying. This allows for a "dry run" to see folder structure and counts without performing the full export.
if ($WhatIf) {
    Write-Log "WHATIF mode enabled: no folders/items will be copied." "WARN"
}
else {
    $pstRoot = Ensure-PstStoreAndRoot -Session $session -PstPath $pstPath
    Write-Log "Beginning folder/item copy from Public Folder to PST..."
    Copy-FolderRecursiveWithStats -SourceFolder $srcFolder -DestParentFolder $pstRoot -SourceFolderPath $PublicFolderPath -WhatIf:$WhatIf
    $desktop = [Environment]::GetFolderPath('Desktop')
    $countsCsv = Join-Path $desktop ("PF_ExportCounts_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
    $script:ExportStats | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $countsCsv
    Write-Host "Export counts CSV saved: $countsCsv"
    Write-Log "Done. PST created at: $pstPath"
}
