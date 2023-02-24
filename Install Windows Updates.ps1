function Install-WACWUWindowsUpdates {
<#

.SYNOPSIS
Create a scheduled task to run a powershell script file to installs all available windows updates through ComObject, restart the machine if needed.

.DESCRIPTION
Create a scheduled task to run a powershell script file to installs given windows updates through ComObject, restart the machine if needed.
This is a workaround since CreateUpdateDownloader() and CreateUpdateInstaller() methods can't be called from a remote computer - E_ACCESSDENIED.
More details see https://msdn.microsoft.com/en-us/library/windows/desktop/aa387288(v=vs.85).aspx

.ROLE
Administrators

.PARAMETER restartTime
  The user-defined time to restart after update (Optional).

.PARAMETER updateIDs
  the list of update IDs to be installed

#>

param (
  [Parameter(Mandatory = $false)]
  [String]$restartTime,
  [Parameter(Mandatory = $false)]
  [Boolean]$skipRestart,
  [Parameter(Mandatory = $true)]
  [String[]]$updateIDs
)

function installWindowsUpdates() {
  param (
    [String]
    $restartTime,
    [Boolean]
    $skipRestart,
    [String[]]
    $updateIDs
  )
  $objServiceManager = New-Object -ComObject 'Microsoft.Update.ServiceManager';
  $objSession = New-Object -ComObject 'Microsoft.Update.Session';
  $objSearcher = $objSession.CreateUpdateSearcher();
  $serviceName = 'Windows Update';
  $search = 'IsInstalled = 0';
  $objResults = $objSearcher.Search($search);
  $Updates = $objResults.Updates;
  $FoundUpdatesToDownload = $Updates.Count;

  $NumberOfUpdate = 1;
  $objCollectionDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl';
  $updateCount = $updateIDs.Count;
  Foreach ($Update in $Updates) {
    If ($Update.Identity.UpdateID -in $updateIDs) {
      Write-Progress -Activity 'Downloading updates' -Status `"[$NumberOfUpdate/$updateCount]` $($Update.Title)`" -PercentComplete ([int]($NumberOfUpdate / $updateCount * 100));
      $NumberOfUpdate++;
      Write-Debug `"Show` update` to` download:` $($Update.Title)`" ;
      Write-Debug 'Accept Eula';
      $Update.AcceptEula();
      Write-Debug 'Send update to download collection';
      $objCollectionTmp = New-Object -ComObject 'Microsoft.Update.UpdateColl';
      $objCollectionTmp.Add($Update) | Out-Null;

      $Downloader = $objSession.CreateUpdateDownloader();
      $Downloader.Updates = $objCollectionTmp;
      Try {
        Write-Debug 'Try download update';
        $DownloadResult = $Downloader.Download();
      } <#End Try#>
      Catch {
        If ($_ -match 'HRESULT: 0x80240044') {
          Write-Warning 'Your security policy do not allow a non-administator identity to perform this task';
        } <#End If $_ -match 'HRESULT: 0x80240044'#>

        Return
      } <#End Catch#>

      Write-Debug 'Check ResultCode';
      Switch -exact ($DownloadResult.ResultCode) {
        0 { $Status = 'NotStarted'; }
        1 { $Status = 'InProgress'; }
        2 { $Status = 'Downloaded'; }
        3 { $Status = 'DownloadedWithErrors'; }
        4 { $Status = 'Failed'; }
        5 { $Status = 'Aborted'; }
      } <#End Switch#>

      If ($DownloadResult.ResultCode -eq 2) {
        Write-Debug 'Downloaded then send update to next stage';
        $objCollectionDownload.Add($Update) | Out-Null;
      } <#End If $DownloadResult.ResultCode -eq 2#>
    }
  }

  $ReadyUpdatesToInstall = $objCollectionDownload.count;
  Write-Verbose `"Downloaded` [$ReadyUpdatesToInstall]` Updates` to` Install`" ;
  If ($ReadyUpdatesToInstall -eq 0) {
    Return;
  } <#End If $ReadyUpdatesToInstall -eq 0#>

  $NeedsReboot = $false;
  $NumberOfUpdate = 1;

  <#install updates#>
  Foreach ($Update in $objCollectionDownload) {
    Write-Progress -Activity 'Installing updates' -Status `"[$NumberOfUpdate/$ReadyUpdatesToInstall]` $($Update.Title)`" -PercentComplete ([int]($NumberOfUpdate / $ReadyUpdatesToInstall * 100));
    Write-Debug 'Show update to install: $($Update.Title)';

    Write-Debug 'Send update to install collection';
    $objCollectionTmp = New-Object -ComObject 'Microsoft.Update.UpdateColl';
    $objCollectionTmp.Add($Update) | Out-Null;

    $objInstaller = $objSession.CreateUpdateInstaller();
    $objInstaller.Updates = $objCollectionTmp;

    Try {
      Write-Debug 'Try install update';
      $InstallResult = $objInstaller.Install();
    } <#End Try#>
    Catch {
      If ($_ -match 'HRESULT: 0x80240044') {
        Write-Warning 'Your security policy do not allow a non-administator identity to perform this task';
      } <#End If $_ -match 'HRESULT: 0x80240044'#>

      Return;
    } #End Catch

    If (!$NeedsReboot) {
      Write-Debug 'Set instalation status RebootRequired';
      $NeedsReboot = $installResult.RebootRequired;
    } <#End If !$NeedsReboot#>
    $NumberOfUpdate++;
  } <#End Foreach $Update in $objCollectionDownload#>

  If ($NeedsReboot -and -not $skipRestart) {
    <#Restart almost immediately, given some seconds for this PSSession to complete.#>
    $waitTime = 5
    if ($restartTime) {
      <#Restart at given time#>
      $waitTime = [decimal]::round(((Get-Date $restartTime) - (Get-Date)).TotalSeconds);
      if ($waitTime -lt 5 ) {
        $waitTime = 5
      }
    }
    Shutdown -r -t $waitTime -c "SME installing Windows updates";
  }
}

#---- Script execution starts here ----

#a dynamic script is created which is executed through the task Scheduler
$ScriptFile = $env:LocalAppData + "\Install-Updates.ps1"

$HashArguments = @{};
if ($restartTime) {
    $HashArguments.Add("restartTime", $restartTime)
}
if ($skipRestart -ne $null) {
    $HashArguments.Add("skipRestart", $skipRestart)
}

$tempArgs = ""
foreach ($key in $HashArguments.Keys) {
    $value = $HashArguments[$key]
    if ($value.GetType().Name -eq "String") {
      $value = """$value"""
    }
    elseif ($value.GetType().Name -eq "Boolean") {
      $value = if ($value -eq $true) { "`$true" } else { "`$false" }
    }
    $tempArgs += " -$key $value"
}

#Create a scheduled task
$TaskName = "SMEWindowsUpdateInstallUpdates"
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

#$OFS is a special variable that contains the string to be used as the Ouptut Field Separator.
#This string is used when an array is converted to a string.  By default, this is " " (white space).
#Change it to separate string array $updateIDs as xxxxx,yyyyyy,zzzzz etc.
$OFS = ","
$tempUpdateIds = [string]$updateIDs


(Get-Command installWindowsUpdates).ScriptBlock | Set-Content -path $ScriptFile
$arg = "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -Command ""&{Set-Location -Path $env:LocalAppData; .\Install-Updates.ps1 $tempArgs -updateIDs $tempUpdateIds }"""

if (!$Role) {
  Write-Warning "To perform some operations you must run an elevated Windows PowerShell console."
}

$Scheduler = New-Object -ComObject Schedule.Service

#Try to connect to schedule service 3 time since it may fail the first time
for ($i = 1; $i -le 3; $i++) {
  Try {
    $Scheduler.Connect()
    Break
  }
  Catch {
    if ($i -ge 3) {
      Write-EventLog -LogName Application -Source "SME Windows Updates Install Updates" -EntryType Error -EventID 1 -Message "Can't connect to Schedule service"
      Write-Error "Can't connect to Schedule service" -ErrorAction Stop
    }
    else {
      Start-Sleep -s 1
    }
  }
}

$RootFolder = $Scheduler.GetFolder("\")
#Delete existing task
if ($RootFolder.GetTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Write-Debug("Deleting existing task" + $TaskName)
  $RootFolder.DeleteTask($TaskName, 0)
}

$Task = $Scheduler.NewTask(0)
$RegistrationInfo = $Task.RegistrationInfo
$RegistrationInfo.Description = $TaskName
$RegistrationInfo.Author = $User.Name

$Triggers = $Task.Triggers
$Trigger = $Triggers.Create(7) #TASK_TRIGGER_REGISTRATION: Starts the task when the task is registered.
$Trigger.Enabled = $true

$Settings = $Task.Settings
$Settings.Enabled = $True
$Settings.StartWhenAvailable = $True
$Settings.Hidden = $False

$Action = $Task.Actions.Create(0)
$Action.Path = "powershell"
$Action.Arguments = $arg

#Tasks will be run with the highest privileges
$Task.Principal.RunLevel = 1

#Start the task to run in Local System account. 6: TASK_CREATE_OR_UPDATE
$RootFolder.RegisterTaskDefinition($TaskName, $Task, 6, "SYSTEM", $Null, 1) | Out-Null
#Wait for running task finished
$RootFolder.GetTask($TaskName).Run(0) | Out-Null
while ($Scheduler.GetRunningTasks(0) | Where-Object { $_.Name -eq $TaskName }) {
  Start-Sleep -s 1
}

#Clean up
$RootFolder.DeleteTask($TaskName, 0)
Remove-Item $ScriptFile


}

Install-WACWUWindowsUpdates