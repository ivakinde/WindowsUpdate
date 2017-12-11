#Windows Updates CLI
 
$UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
$UpdateSession.ClientApplicationID = 'PowerShell Sample'
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
 
Write-Host 'Searching for updates...' -ForegroundColor Green
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
 
if ($SearchResult.Updates.Count -ne 0)
{
Write-Host 'There are ' $SearchResult.Updates.Count 'applicable updates on the machine:' -ForegroundColor Green
}
else
{
Write-Host 'There are no applicable updates' -ForegroundColor Green
break
}
Write-Host 'Creating a collection of updates to download:' -ForegroundColor Green
$UpdatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
foreach ($Update in $SearchResult.Updates)
{
[bool]$addThisUpdate = $false
if ($Update.InstallationBehavior.CanRequestUserInput)
{
Write-Host "> Skipping: $($Update.Title) because it requires user input" -ForegroundColor Green
}
else
{
if (!($Update.EulaAccepted))
{
Write-Host "> Note: $($Update.Title) has a license agreement that must be accepted:"
$Update.EulaText
$strInput = Read-Host 'Do you want to accept this license agreement? (Y/N)'
if ($strInput.ToLower() -eq 'y')
{
$Update.AcceptEula()
[bool]$addThisUpdate = $true
}
else
{
Write-Host "> Skipping: $($Update.Title) because the license agreement was declined"
}
}
else
{
[bool]$addThisUpdate = $true
}
}
if ([bool]$addThisUpdate)
{
Write-Host "Adding: $($Update.Title)"
$UpdatesToDownload.Add($Update) |Out-Null
}
}
 
if ($UpdatesToDownload.Count -eq 0)
{
Write-Host 'All applicable updates were skipped.' -ForegroundColor Green
break
}
 
Write-Host ""
Write-Host 'Downloading updates...' -ForegroundColor Green
$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()
 
$UpdatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
 
[bool]$rebootMayBeRequired = $false
Write-Host 'Successfully downloaded updates:' -ForegroundColor Green
 
foreach ($Update in $SearchResult.Updates)
{
if ($Update.IsDownloaded)
{
Write-Host "> $($Update.Title)"
$UpdatesToInstall.Add($Update)
 
if ($Update.InstallationBehavior.RebootBehavior -gt 0)
{
[bool]$rebootMayBeRequired = $true
}
}
}
 
if ($UpdatesToInstall.Count -eq 0)
{
Write-Host 'No updates were succsesfully downloaded' -ForegroundColor Green
}
 
if ($rebootMayBeRequired)
{
Write-Host 'These updates may require a reboot' -ForegroundColor Green
}
 
$strInput = Read-Host "Would you like to install updates now? (Y/N)"
 
if ($strInput.ToLower() -eq 'y')
{
Write-Host 'Installing updates...' -ForegroundColor Green
 
$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallationResult = $Installer.Install()
 
Write-Host "Installation Result: $($InstallationResult.ResultCode)" -ForegroundColor Green
Write-Host "Reboot Required: $($InstallationResult.RebootRequired)" -ForegroundColor Green
Write-Host 'Listing of updates installed and individual installation results' -ForegroundColor Green
 
for($i=0; $i -lt $UpdatesToInstall.Count; $i++)
{
Write-Host "> $($Update.Title) : $($InstallationResult.GetUpdateResult($i).ResultCode)"
}
}
