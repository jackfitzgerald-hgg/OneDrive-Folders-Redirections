<#

Function BackupNetworkShortcuts() {
	[Helios]::WriteLog("Backing up network shortcuts")
	# Check 7Zip exists
	If (!(Test-Path $7Zip)) {
		[Helios]::WriteLog("Unable to find ""$7Zip""")
		Return
	}
	# Check network shortcuts folder exists
	$SourceFolder = $Env:AppData + "\Microsoft\Windows\Network Shortcuts"
	If (!(Test-Path $SourceFolder)) {
		[Helios]::WriteLog("Unable to find ""$SourceFolder""")
		Return
	}
	# Archive network shortcuts
	$Now = Get-Date
	$Suffix = "{0:0000}-{1:00}-{2:00}-{3:00}-{4:00}" -f $Now.Year, $Now.Month, $Now.Day, $Now.Hour, $Now.Minute
	$ArchivePath = "$NetworkShortcutsBackupFolder\$($Env:ComputerName) - $Suffix.7z"
	$Cmd = "& '$7Zip' a -r '$ArchivePath' '$SourceFolder'"
	[void](Invoke-Expression $Cmd)
	# Remove old backups
	$KeepDate = (Get-Date).AddDays(-$NetworkShortcutsRetention)
	Get-ChildItem -Path $NetworkShortcutsBackupFolder -File | Where {$_.LastWriteTime -le $KeepDate} | Remove-Item
	# Remove old format backups
	Get-ChildItem -Path $NetworkShortcutsBackupFolder -Exclude *.7z | Remove-Item -Recurse -Force
}


#>


<#

Backing up Network Shortcuts
Jack Fitzgerald
Helios Global Group

03/22

#>

# Variables

$OneDriveFolder = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\Business1' -Name 'UserFolder'
$7Zip = "$([Environment]::GetFolderPath("ProgramFiles"))\7-Zip\7z.exe"      # 7Zip Path
$NetworkShortcutsBackupFolder = "$OneDriveFolder\Backup\Network Shortcuts"		# Folder for network locations backup
$NetworkShortcutsRetention = 60												# Number of days to keep network shortcut backups


# Check 7Zip exists

if(!(Test-Path $7Zip)){
    exit
    Write-Host "No 7Zip"
}

# Check network shortcuts folder exists

$SourceFolder = $Env:AppData + "\Microsoft\Windows\Network Shortcuts"
	if (!(Test-Path $SourceFolder)) {
        Write-Host "No Folder Found"
        exit
	}

# Archive network shortcuts

$Now = Get-Date
$Suffix = "{0:0000}-{1:00}-{2:00}-{3:00}-{4:00}" -f $Now.Year, $Now.Month, $Now.Day, $Now.Hour, $Now.Minute
$ArchivePath = "$NetworkShortcutsBackupFolder\$($Env:ComputerName) - $Suffix.7z"
$Cmd = "& '$7Zip' a -r '$ArchivePath' '$SourceFolder'"
[void](Invoke-Expression $Cmd)

# Remove old backups

$KeepDate = (Get-Date).AddDays(-$NetworkShortcutsRetention)
Get-ChildItem -Path $NetworkShortcutsBackupFolder -File | Where {$_.LastWriteTime -le $KeepDate} | Remove-Item

# Remove old format backups

Get-ChildItem -Path $NetworkShortcutsBackupFolder -Exclude *.7z | Remove-Item -Recurse -Force