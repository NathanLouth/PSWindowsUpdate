function Get-WACWUAvailableWindowsUpdates {
<#

.SYNOPSIS
Get available windows updates through COM object by Windows Update Agent API.

.DESCRIPTION
Get available windows updates through COM object by Windows Update Agent API.

.ROLE
Readers

#>

$objSession = Microsoft.PowerShell.Utility\New-Object -ComObject "Microsoft.Update.Session"
$objSearcher = $objSession.CreateUpdateSearcher()
$objResults = $objSearcher.Search("IsInstalled = 0")

if (!$objResults -or !$objResults.Updates) {
  return $null
}

<#
InstallationBehavior.RebootBehaviour enum
	0: NeverReboots
	1: AlwaysRequiresReboot
  2: CanRequestReboot

InstallationBehavior.Impact enum
  0: Normal
	1: Minor
	2: RequiresExclusiveHandling
#>
$objResults.Updates | ForEach-Object {
  New-Object PSObject -Property @{
    Title                       = $_.Title
    IsMandatory                 = $_.IsMandatory
    RebootRequired              = $_.RebootRequired
    MsrcSeverity                = $_.MsrcSeverity
    IsUninstallable             = $_.IsUninstallable
    UpdateID                    = ($_ | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty Identity).UpdateID
    KBArticleIDs                = $_ | Microsoft.PowerShell.Utility\Select-Object  KBArticleIDs | ForEach-Object { $_.KbArticleids }
    CanRequestUserInput         = ($_ | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty InstallationBehavior).CanRequestUserInput
    Impact                      = ($_ | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty InstallationBehavior).Impact
    RebootBehavior              = ($_ | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty InstallationBehavior).RebootBehavior
    RequiresNetworkConnectivity = ($_ | Microsoft.PowerShell.Utility\Select-Object -ExpandProperty InstallationBehavior).RequiresNetworkConnectivity
  }
}

}

Get-WACWUAvailableWindowsUpdates