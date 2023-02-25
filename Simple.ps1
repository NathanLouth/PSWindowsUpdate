# Create a new COM object for the Windows Update Agent API
$updateSession = New-Object -ComObject Microsoft.Update.Session

# Get a collection of all available updates
$updateSearcher = $updateSession.CreateUpdateSearcher()
$updateSearchResult = $updateSearcher.Search("IsInstalled=0")

# Print the list of updates to the console
foreach ($update in $updateSearchResult.Updates) {
    Write-Output "Title: $($update.Title)"
    Write-Output "Description: $($update.Description)"
    Write-Output "Downloaded: $($update.IsDownloaded)"
    Write-Output "Mandatory: $($update.IsMandatory)"
    Write-Output "Update Type: $($update.Type)"
    Write-Output "Publication State: $($update.PublicationState)"
    Write-Output "Creation Date: $($update.CreationDate)"
    Write-Output "Last Deployment Change Time: $($update.LastDeploymentChangeTime)"
    Write-Output "Languages: $($update.Languages)"
    Write-Output "Products: $($update.Products)"
    $updateProperties = $update.Properties
    $rebootRequired = $updateProperties | Where-Object { $_.Name -eq "RebootRequired" } | Select-Object Value
    Write-Output "Reboot Required: $($rebootRequired.Value)"
    Write-Output "-----"
}

# Prompt the user to install all available updates
$installAllUpdates = Read-Host "Do you want to install all available updates? (Y/N)"

if ($installAllUpdates -eq "Y") {
    # Download all updates
    $updateDownloader = $updateSession.CreateUpdateDownloader()
    $updateDownloader.Updates.AddRange($updateSearchResult.Updates)
    $downloadResult = $updateDownloader.Download()

    # Install all updates
    $updateInstaller = $updateSession.CreateUpdateInstaller()
    $updateInstaller.Updates.AddRange($updateSearchResult.Updates)
    $installResult = $updateInstaller.Install()

    if ($downloadResult.HResult -eq 0 -and $installResult.HResult -eq 0) {
        Write-Output "All updates successfully installed."
    } else {
        Write-Output "Error installing updates."
    }
}

# Check if a reboot is required
$systemRebootRequired = $updateSession.SystemRebootRequired

if ($systemRebootRequired) {
    Write-Output "A system reboot is required to complete the installation of updates."
} else {
    Write-Output "No system reboot is required."
}
