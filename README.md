# ExportPublicFolder
Export a public folder to a pst file using outlook mapi profile via powershell. Exports Calendars, Mail items, and Tasks Folder Type and maintains folder structure.

Written in VSCode with assistance from GitHub Copilot

# Requirements

- Ensure path to EWS dll file exists (requires download of EWS Managed API 2.2 via NuGet https://www.nuget.org/packages/Microsoft.Exchange.WebServices)
- Outlook (classic mode) open using a Profile with R/W to the public folder to be exported
- PST file will be created on the desktop. Edit variable $pstFileName OR set filename using the -pstFileName switch in the powershell command


# Usage
From Powershell running as standard user (Not Administrator):

 .\exportPublicFolder.ps1 -PublicFolderPath "\PF-HumanResources" -WhatIf
 
# generates a preview of the public folder's folder structure and item counts

# remove the -WhatIf switch to perform export to pst
