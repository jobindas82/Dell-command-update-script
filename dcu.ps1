if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -ExecutionPolicy RemoteSigned -File `"$PSCommandPath`"" -Verb RunAs; 
	exit 
}

Clear-Host
Write-Host " ___   ___ _   _         ___      _               ___         _        _ _         " -ForegroundColor green
Write-Host "|   \ / __| | | |  ___  |   \ _ _(_)_ _____ _ _  |_ _|_ _  __| |_ __ _| | |___ _ _ " -ForegroundColor green
Write-Host "| |) | (__| |_| | |___| | |) | '_| \ V / -_) '_|  | || ' \(_-<  _/ _` | | / -_) '_|" -ForegroundColor green
Write-Host "|___/ \___|\___/        |___/|_| |_|\_/\___|_|   |___|_||_/__/\__\__,_|_|_\___|_|   "  -ForegroundColor green
Write-Host " " -ForegroundColor green

$FilesRoot = $PSScriptRoot + "\files"
$Model = (Get-CimInstance Win32_ComputerSystemProduct | Select-Object Name).Name

$CatalogUrl = "https://dl.dell.com/catalog/DriverPackCatalog.cab"
$DriverURL = "https://dl.dell.com"
$CatalogCABFile = "$FilesRoot\catelog.cab"
$CatalogXMLFile = "$FilesRoot\catelog.xml"
$versionFile = "$FilesRoot\versions.xml"
$versionFileBackup = "$FilesRoot\versions.backup.xml"
$dcuCliPath = "C:\Program Files\Dell\CommandUpdate"

if (Test-Path "$dcuCliPath\dcu-cli.exe") {}
else {
	$dcuCliPath = "C:\Program Files (x86)\Dell\CommandUpdate"
}

#update here
$TargetModels = "Latitude 5440 Latitude 5430 Latitude 5420 Latitude 5410 Latitude 5400 Precision 5690 Precision 5680 Precision 3660 Tower Precision 3650 Tower Precision 5540 Precision 5550 Precision 5560 Precision 5570"

#Update here
$TargetOS = "Windows 10 x64 Windows 11 x64"

if ((Get-CimInstance Win32_operatingsystem).OSArchitecture -eq "64-Bit") {
	$Arch = "_x64"
}
elseif ( (Get-CimInstance Win32_operatingsystem).OSArchitecture -eq "32-Bit") {
	$Arch = "_x86"
}
else {
	Write-Host "Architecture not supported."
	Exit
}

if (((Get-WmiObject Win32_OperatingSystem).Caption).Contains("Windows 10")) {
	$OsName = "Windows_10"
}
elseif (((Get-WmiObject Win32_OperatingSystem).Caption).Contains("Windows 11")) {
	$OsName = "Windows_11"
}
else {
	Write-Host "OS not supported."
	Exit
}

function Invoke-Automation {
	Remove-Dcu
	Install-Dcu
	Set-Driverpack
	Invoke-Restore
	Copy-Update
	Write-Output "Rebooting in 10 Seconds.."
	Start-Sleep -Seconds 10
	Restart-Computer -Force
}

function Test-Dcu {
	$Name = "Dell Command | Update*"
	$ProgramList = @( "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" )
	$Programs = Get-ItemProperty $ProgramList -EA 0
	$App = ($Programs | Where-Object { $_.DisplayName -like $Name -and $_.UninstallString -like "*msiexec*" }).PSChildName

	if ($App) {
	}
	else {
		Install-Dcu
	}
}

function Install-Dcu {
	if (Test-Path "$FilesRoot\dcu\setup.exe") {
		Write-Output "Installing the latest dell command update..."
		& "$FilesRoot\dcu\setup.exe" /s  | Out-Null
	}
	else {
		Write-Output "Dell command installer not found.."
		Start-Exec
	}


	$dcuCliPath = "C:\Program Files\Dell\CommandUpdate"
	if (Test-Path "$dcuCliPath\dcu-cli.exe") {}
	else {
		$dcuCliPath = "C:\Program Files (x86)\Dell\CommandUpdate"
	}
}

function Set-Driverpack {
	& "$dcuCliPath\dcu-cli.exe" /configure -advancedDriverRestore=enable
	Write-Output "`r`Enabled Advanced Driver Restore..."
	if (Test-Path "$FilesRoot\$OsName$Arch\$($Model.Replace(' ', '_').Replace('_Tower', ''))\pack.exe") {
		Write-Output "Using offline driver pack.."
		& "$dcuCliPath\dcu-cli.exe" /configure -driverLibraryLocation="$FilesRoot\$OsName$Arch\$($Model.Replace(' ', '_').Replace('_Tower', ''))\pack.exe" | Out-Null
	}
	elseif (Test-Path "$FilesRoot\$OsName$Arch\$($Model.Replace(' ', '_').Replace('_Tower', ''))\pack.cab") {
		Write-Output "Using offline driver pack.."
		& "$dcuCliPath\dcu-cli.exe" /configure -driverLibraryLocation="$FilesRoot\$OsName$Arch\$($Model.Replace(' ', '_').Replace('_Tower', ''))\pack.cab" | Out-Null
	}
	else {
		while (!(Test-Connection -ComputerName www.google.com -Quiet)) {
			Write-Host "Please check your internet connection.. retrying in 10 Seconds.."
			Start-Sleep -Seconds 10
		}
	}
}

function Invoke-Restore {
	& "$dcuCliPath\dcu-cli.exe" /driverinstall
	& "$dcuCliPath\dcu-cli.exe" /configure -advancedDriverRestore=disable
}

function Remove-Dcu {
	Write-Output "Removing any existing dell command update installations..."
	$Name = "Dell Command | Update*"
	$ProcName = "DellCommandUpdate"
	$Timestamp = Get-Date -Format "yyyy-MM-dd_THHmmss"
	$LogFile = "$env:TEMP\Dell-CU-Uninst_$Timestamp.log"
	$ProgramList = @( "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" )
	$Programs = Get-ItemProperty $ProgramList -EA 0
	$App = ($Programs | Where-Object { $_.DisplayName -like $Name -and $_.UninstallString -like "*msiexec*" }).PSChildName

	if ($App) {
		Get-Process | Where-Object { $_.ProcessName -eq $ProcName } | Stop-Process -Force
		$Params = @(
			"/qn"
			"/norestart"
			"/X"
			"$App"
			"/L*V ""$LogFile"""
		)
		Start-Process "msiexec.exe" -ArgumentList $Params -Wait 
	}
	else {
		Write-Output "$Name not found installed in registry."

	}
}

function Copy-Update {

	$psContent = @"
clear
Write-Host " ___   ___ _   _         ___      _               ___         _        _ _         " -ForegroundColor green
Write-Host "|   \ / __| | | |  ___  |   \ _ _(_)_ _____ _ _  |_ _|_ _  __| |_ __ _| | |___ _ _ " -ForegroundColor green
Write-Host "| |) | (__| |_| | |___| | |) | '_| \ V / -_) '_|  | || ' \(_-<  _/ _` | | / -_) '_|" -ForegroundColor green
Write-Host "|___/ \___|\___/        |___/|_| |_|\_/\___|_|   |___|_||_/__/\__\__,_|_|_\___|_|   "  -ForegroundColor green
Write-Host " " -ForegroundColor green                                                                           

while(!(Test-Connection -ComputerName www.google.com -Quiet)){
	Write-Host "Please check your internet connection.. retrying in 10 Seconds.."
	Start-Sleep -Seconds 10
}

Write-Host "Process will start in 5 minutes.."
Start-Sleep -Seconds 300

`$retry` = 1;
while(`$retry` -le 5){
	& "$dcuCliPath\dcu-cli.exe" /scan
	Start-Sleep -Seconds 10
	& "$dcuCliPath\dcu-cli.exe" /ApplyUpdates
	if((`$LastExitCode` -eq "501") -Or (`$LastExitCode` -eq "502") -Or (`$LastExitCode` -eq "503") -Or (`$LastExitCode` -eq "1000") -Or (`$LastExitCode` -eq "1001") -Or (`$LastExitCode` -eq "1002")){
		Write-Host "`r`Retrying again in 2 minutes.." 
		Start-Sleep -Seconds 120
		`$script:retry++`
	}else{
		Break
	}
}
Write-Host "`r`Operation completed..." 
if(Test-Path "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\dcu.cmd")
{
	Remove-Item "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup\dcu.cmd"
}
Start-Sleep -Seconds 5
Exit
"@

	New-Item -Path ([Environment]::GetFolderPath("Desktop")) -Name "dcu.ps1" -ItemType "file" -Value $psContent -Force | Out-Null

	$cmdContent = @"
powershell.exe "& '$(([Environment]::GetFolderPath("Desktop")))\dcu.ps1'"
"@

	New-Item -Path "$($env:APPDATA)\Microsoft\Windows\Start Menu\Programs\Startup" -Name "dcu.cmd" -ItemType "file" -Value $cmdContent -Force | Out-Null
}

function Show-Model {
	Write-Host "Asset standard   : $Model" -ForegroundColor red
	Write-Host "Windows standard : $($OsName.Replace('_', ' '))$($Arch.Replace('_', ' '))" -ForegroundColor red
}

function Add-Root-Folder {
	param (
		[string]$FolderName
	)
	if (!(Test-Path $FolderName)) {
		Try {
			New-Item -Path $FolderName -ItemType Directory -Force | Out-Null
		}
		Catch {
			Write-Error "$($_.Exception)"
		}
	}
}

function Remove-Catelog {
	if (Test-Path "$FilesRoot\catelog.cab") {
		Remove-Item -Path "$FilesRoot\catelog.cab" -Force | Out-Null
	}
	if (Test-Path "$FilesRoot\catelog.xml") {
		Remove-Item -Path "$FilesRoot\catelog.xml" -Force | Out-Null
	}
}

function convertFileSize {
	param(
		$bytes
	)

	if ($bytes -lt 1MB) {
		return "$([Math]::Round($bytes / 1KB, 2)) KB"
	}
	elseif ($bytes -lt 1GB) {
		return "$([Math]::Round($bytes / 1MB, 2)) MB"
	}
	elseif ($bytes -lt 1TB) {
		return "$([Math]::Round($bytes / 1GB, 2)) GB"
	}
}
	
function Get-Catelog {
	Remove-Catelog
	$wc = New-Object System.Net.WebClient
	$wc.Headers.Add([System.Net.HttpRequestHeader]::UserAgent, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
	$wc.DownloadFile($CatalogUrl, "$FilesRoot\catelog.cab")
	if (!(Test-Path "$FilesRoot\catelog.cab")) {
		Write-Error "Download Failed. Exiting Script."
	}
	$wc.Dispose()
}
	
function Get-Driver-Pack {
	Add-Root-Folder($FilesRoot)
	Get-Catelog

	Write-Host "Checking for updates from dell..."
	EXPAND $CatalogCABFile $CatalogXMLFile | Out-Null

	[XML]$Catalog = Get-Content $CatalogXMLFile
	[array]$DriverPackages = $Catalog.DriverPackManifest.DriverPackage

	foreach ($DriverPackage in $DriverPackages) {
		
		$DriverPackageDownloadPath = "$DriverURL/$($DriverPackage.path)"
		$DriverPackageName = $DriverPackage.Name.Display.'#cdata-section'.Trim()
		
		if (!$DriverPackage.SupportedSystems) {
			continue
		}
		
		foreach ($Brand in $DriverPackage.SupportedSystems.Brand) {
			$Model = $Brand.Model.name.Trim()
			break
		}
		
		if (!$TargetModels.Contains($Model)) {
			continue
		}
		
		foreach ($SupportedOS in $DriverPackage.SupportedOperatingSystems) {
			if (!$TargetOS.Contains($SupportedOS.OperatingSystem.Display.'#cdata-section'.Trim())) {
				continue
			}
			
			$newDriverFile = $true
			$OsSupported = $SupportedOS.OperatingSystem.Display.'#cdata-section'.Trim().Replace(' ', '_')
			$DownloadDestination = "$FilesRoot\$OsSupported\$($Model.Replace(' ', '_').Replace('_Tower', ''))"

			if (Test-Path "$versionFile") {
				Copy-Item $versionFile -Destination $versionFileBackup -Force
				[XML]$VersionContent = Get-Content $versionFile

				foreach ($fileVersion in $VersionContent.Models.Model) {
					if ( ($fileVersion.name -eq $Model.Replace(' ', '_')) -and ($fileVersion.releaseID -eq $DriverPackage.releaseID) -and ($fileVersion.os -eq $OsSupported ) ) {
						
						if (!(Test-Path "$DownloadDestination\pack.cab") -and !(Test-Path "$DownloadDestination\pack.exe") ) {
							$newDriverFile = $true
							Break
						}
						
						if (Test-Path "$DownloadDestination\pack.exe") {
							$driverFile = Get-Item "$DownloadDestination\pack.exe"
							$SizeinGB = $driverFile.Length / 1GB
							if ($SizeinGB -le 1) {
								$newDriverFile = $true
								Break
							}
						}
						
						if (Test-Path "$DownloadDestination\pack.cab") {
							$driverFile = Get-Item "$DownloadDestination\pack.cab"
							$SizeinGB = $driverFile.Length / 1GB
							if ($SizeinGB -le 1) {
								$newDriverFile = $true
								Break
							}
						}

						$newDriverFile = $false
						Break
					}
				}
			}

			if (!$newDriverFile) {
				continue
			}

			while ($true) {
				if (Test-Path "$DownloadDestination") {
					Remove-Item -Path "$DownloadDestination" -Force -Recurse | Out-Null
				}
				Add-Root-Folder($DownloadDestination)
			
				$webClient = New-Object -TypeName System.Net.WebClient
				$webClient.Headers.Add([System.Net.HttpRequestHeader]::UserAgent, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
				$task = $webClient.DownloadFileTaskAsync($DriverPackageDownloadPath, "$DownloadDestination\$DriverPackageName")

				Register-ObjectEvent -InputObject $webClient -EventName DownloadProgressChanged -SourceIdentifier WebClient.DownloadProgressChanged | Out-Null

				Start-Sleep -Seconds 3

				while (!($task.IsCompleted)) {
					$EventData = Get-Event -SourceIdentifier WebClient.DownloadProgressChanged | Select-Object -ExpandProperty "SourceEventArgs" -Last 1

					$ReceivedData = ($EventData | Select-Object -ExpandProperty "BytesReceived")
					$TotalToReceive = ($EventData | Select-Object -ExpandProperty "TotalBytesToReceive")
					$TotalPercent = $EventData | Select-Object -ExpandProperty "ProgressPercentage"

					Start-Sleep -Seconds 2

					Write-Progress -Activity "Downloading $DriverPackageName" -Status "Percent Complete: $($TotalPercent)%" -CurrentOperation "Downloaded $(convertFileSize -bytes $ReceivedData) / $(convertFileSize -bytes $TotalToReceive)" -PercentComplete $TotalPercent
				}

				Unregister-Event -SourceIdentifier WebClient.DownloadProgressChanged
				$webClient.Dispose()
				
				$driverFile = Get-Item "$DownloadDestination\$DriverPackageName"
				$SizeinGB = $driverFile.Length / 1GB
				if ($SizeinGB -le 1) {
					Write-Host "Download failed: Restarting download.."
				}
				else {
					Get-ChildItem -Path $DownloadDestination -Filter * | Rename-Item -NewName { "pack" + $_.extension }
					if (!(Test-Path "$versionFile")) {
						$xmlsettings = New-Object System.Xml.XmlWriterSettings
						$xmlsettings.Indent = $true
						$xmlWriter = [System.XML.XmlWriter]::Create($versionFile, $xmlsettings)
						$xmlWriter.WriteStartElement('Models') 
						$xmlWriter.WriteStartElement('Model') 
						$xmlWriter.WriteAttributeString("name", $Model.Replace(' ', '_'))
						$xmlWriter.WriteAttributeString("releaseID", $DriverPackage.releaseID)
						$xmlWriter.WriteAttributeString("releaseDate", $DriverPackage.dateTime)
						$xmlWriter.WriteAttributeString("os", $OsSupported)
						$xmlWriter.WriteEndElement()
						$xmlWriter.WriteEndElement()
						$xmlWriter.Flush()
						$xmlWriter.Close()

					}
					else {
						$ModelFound = $false
						foreach ($ModelEl in $VersionContent.Models.Model) {
							if ( ($ModelEl.name -eq $Model.Replace(' ', '_')) -and ($ModelEl.os -eq $OsSupported ) ) {
								$ModelFound = $true
								$ModelEl.releaseID = $DriverPackage.releaseID
								$ModelEl.releaseDate = $DriverPackage.dateTime
								$VersionContent.Save($versionFile)
								Break
							}
						}
						if (!$ModelFound) {
							$newModel = $VersionContent.CreateElement("Model")
							$newModelEl = $VersionContent.Models.AppendChild($newModel)
							$newModelEl.SetAttribute("name", $Model.Replace(' ', '_'))
							$newModelEl.SetAttribute("releaseID", $DriverPackage.releaseID)
							$newModelEl.SetAttribute("releaseDate", $DriverPackage.dateTime)
							$newModelEl.SetAttribute("os", $OsSupported)
							$VersionContent.Save($versionFile)
						}
					}
					Write-Host "Completed: $DriverPackageName"
					Break
				}
			}				
		}
	}
}

function Invoke-Advanced-Restore {
	Show-Model
	Invoke-Automation
}

function Suspend-Teams {
	Get-process ms-teams* | Stop-Process
}

function Start-Exec {
	Suspend-Teams
	Write-Host "  "
	Write-Host "  "
	Write-Host "1. Run Automated Driver restore and updates (Unattended)"
	#Write-Host "2. Run Automated Driver restore and updates (Unattended)"
	Write-Host "3. Install latest drivers from dell online"
	Write-Host "4. Install Dell command update application"
	Write-Host "5. Uninstall dell command update application"
	#Write-Host "6. Install NVIDIA Control Panel (Machine must be added to domain)"
	Write-Host "7. Check and download latest driver packs from Dell website. (For administration only)"
	Write-Host "  "
	$option = Read-Host -Prompt "Please choose one option to continue"

	if ($option -eq "1") {
		tzutil /s "Pacific Standard Time"
		Invoke-Advanced-Restore
	}
	elseif ($option -eq "2") {
		tzutil /s "Mountain Standard Time"
		Invoke-Advanced-Restore
	}
	elseif ($option -eq "4") {
		Remove-Dcu
		Install-Dcu
		Write-Host -Prompt "`r`Installed Dell Command Update application..."
		Start-Exec
	}
	elseif ($option -eq "3") {
		Test-Dcu
		& "$dcuCliPath\dcu-cli.exe" /ApplyUpdates
		Write-Host -Prompt "`r`Installed latest dell drivers..."
		Start-Exec
	}
	elseif ($option -eq "7") {
		Get-Driver-Pack
		Write-Host -Prompt "`r`Checked and updated Driver packs..."
		Start-Exec
	}
	elseif ($option -eq "6") {
		#nvidia
		Start-Exec
	}
	elseif ($option -eq "5") {
		Remove-Dcu
		Write-Host -Prompt "`r`Removed Dell Command Update application..."
		Start-Exec
	}
	else {
		Write-Host -Prompt "`r`Invalid option selected..."
		Start-Exec
	}
}

Start-Exec