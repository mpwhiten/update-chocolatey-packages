<#
Changes:
v3:
Updated log deletion to perform as an array. Much fewer lines of code.
Moved logs to 'Logs' subfolder.

v3.1:
Streamlining the Chocolatey upgrade step at the beginning to only unpin chocolatey in the case an upgrade is required. Should hopefully prevent hanging on this line and speed up execution.
Removed '-Wait' switch on choco process to check if this helps with the pausing between operations.
Added time log for the start of each package upgrade
Added Tee-Object to Chocolatey upgrade output

v3.2:
Change Chocolatey upgrade process from automated to prompt.
Update Chocolatey upgrade check to latest commands.
Make it clearer where each upgrade process begins by spacing out the date print.


Further improvements:
Set up log files as objects
#>

$host.UI.RawUI.WindowTitle = "Update Chocolatey Packages"

# Update choco packages
$chocoPath = "D:\System\Chocolatey\"
Set-Location $chocoPath

$printDate = {
	Write-Host "===================" -ForegroundColor Red -NoNewline; Write-Host ([char]0xA0)
    Write-Host " Date: $(Get-Date -UFormat "%d %b %Y") " -ForegroundColor DarkYellow -NoNewline; Write-Host ([char]0xA0)
    Write-Host " Time: $(Get-Date -f "HH:mm:ss") " -ForegroundColor DarkYellow -NoNewline; Write-Host ([char]0xA0)
	Write-Host "===================" -ForegroundColor Red -NoNewline; Write-Host ([char]0xA0)
}
.$printDate
$i = 0

Write-Host "`n Update Chocolatey Packages " -ForegroundColor Black -BackgroundColor Yellow -NoNewline; Write-Host ([char]0xA0)

$filesDates = {
	$date = Get-Date -Format "-yyyy-MM-dd-HHmm"

	$logDupes = ".\Logs\ChocoDupes$($date).json"
	$logInstalled = ".\Logs\ChocoInstalledPackages$($date).json"
	$logOutdated = ".\Logs\ChocoOutdated$($date).json"
	$logTable = ".\Logs\ChocoOutdatedTable$($date).json"
	$errLog = ".\Logs\ChocoUpdateError$($date).txt"
	$sucsLog = ".\Logs\ChocoUpdateSuccess$($date).txt"
}

Write-Host "`nChecking for latest version of Chocolatey and upgrading..." -ForegroundColor DarkGray
# $chocoCheck = (choco upgrade chocolatey -r --noop).Split('|')
$chocoCur = (choco list chocolatey --limit-output --exact).Split('|')
$chocoCheck = (choco info chocolatey --limit-output).Split('|')
If ($chocoCheck[1] -eq $chocoCur[1]){
	Write-Host "No update required. Version $($chocoCheck[1]) installed is the latest" -ForegroundColor Green
} ElseIf ($chocoCheck[1] -ne $chocoCur[1]){
	Write-Host "Update required." -ForegroundColor Yellow
	Write-Host "Current version: $($chocoCur[1])"
	Write-Host "Latest version: $($chocoCheck[1])"
	
	$upgrade = Read-Host -Prompt "`nInstall latest version now? (y/n)"
	If ($upgrade -like 'y'){
		choco pin remove --name chocolatey | Out-Null
		choco upgrade chocolatey --limit-output | Tee-Object -Variable chocoUp
		choco pin add --name chocolatey | Out-Null
	
		If ($chocoUp -match "Chocolatey upgraded 1/1"){
			Write-Host "Successfully updated Chocolatey to version $($chocoCheck[2])" -ForegroundColor Green
			$chocoUp | Out-File -FilePath $sucsLog
			Write-Host "Chocolatey output saved to $sucsLog" -ForegroundColor DarkGray
		} ElseIf ($chocoUp -match "Chocolatey upgraded 0/1"){
			Write-Host "Upgrade of Chocolatey failed." -ForegroundColor Red
			Write-Host "Printing Chocolatey output and saving to $errLog" -ForegroundColor DarkGray
			$chocoUp | Tee-Object -FilePath $errLog
			Read-Host -Prompt "Press Enter to continue."
		} Else {
			Write-Host "Unknown error occured while updating." -ForegroundColor Red
			$chocoUp | Out-File -FilePath $errLog
			Write-Host "Chocolatey output saved to $errLog" -ForegroundColor DarkGray
			Read-Host -Prompt "Press Enter to continue."
		}
	} Else {
		Write-Host "Skipping upgrade" -ForegroundColor Yellow
	}
}

Do {
	$runAgain = $null
	$i++
	If ($i -gt 1) {"`n"; .$printDate}
	.$filesDates
	Write-Host "`nChecking for outdated packages..." -ForegroundColor DarkGray

	$installed = choco list --local-only --limit-output
	$outdated = choco outdated --limit-output
	$outdated | ConvertTo-Json | Out-File $logOutdated
	# $outdated = $null
	$params = Get-Content .\params.json | ConvertFrom-Json

	$dupes = @()
	ForEach ($line in $installed){
		$package = $line.Split("|")[0]
		If ($package.Split(".").Count -gt 1 -and ($installed -like $($package.Split(".")[0] + "|*"))){
			$dupes += [PSCustomObject] @{
				Duplicate = $line.Split("|")[0]
				BasePackage = ($installed -like $('*' + $line.Split("|")[0].Split(".")[0] + '*'))[0].Split('|')[0]
			}
		}
	}
	
	If ($outdated){
		$listDupes = @()
		$exclude = @()
		$odTable = @()
		ForEach ($line in $outdated){
			$package = $line.Split("|")[0]
			$cVer = $line.Split("|")[1]
			$nVer = $line.Split("|")[2]
			If ($dupes.Duplicate -like $package){
				$listDupes += [array]::IndexOf($dupes.Duplicate,$package)
			} ElseIf ($params.$package.Exclude -eq "Yes"){
				$exclude += [PSCustomObject] @{
					Package = $package
					InstalledVersion = $cVer
					NewVersion = $nVer
					Exclude = $params.$package.Exclude
				}
			} Else {
				$odTable += [PSCustomObject] @{
					Package = $package
					InstalledVersion = $cVer
					NewVersion = $nVer
					Priority = If ($params.$package.Priority){$params.$package.Priority} Else {'1000'}
					Parameters = If ($params.$package.Params){$params.$package.Params} Else {$null}
				}
			}
		}
		
		# Record outdated app data to file
		$odTable | ConvertTo-Json | Out-File $logTable
		# Import most recent table for debugging
		# $odTable = (gci ChocoOutdatedTable* | Sort-Object LastWriteTime)[-1] | Get-Content | ConvertFrom-Json | Sort-Object {[int]($_.Priority -replace '(\d+).*', '$1')},{$_.Package}
		$dupes[$listDupes] | ConvertTo-Json | Out-File $logDupes
		
		If ($odTable.Count -gt 0){
		# If ($odTable.Count -le 0){
			Write-Host "The following outdated packages were found:" -ForegroundColor Green
			$odTable = $odTable | Sort-Object {[int]($_.Priority -replace '(\d+).*', '$1')},{$_.Package}
			$odTable | Format-Table
			
			If ($dupes[$listDupes].Count -gt 0){
				Write-Host "Duplicate packages being ignored:" -ForegroundColor Magenta
				$dupes[$listDupes] | Out-Host
			}
			
			If ($exclude.Count -gt 0){
				Write-Host "Excluded packages being ignored:" -ForegroundColor Magenta
				$exclude | Out-Host
			}
			
			# Read-Host "Press Enter to continue with the upgrade or press CTRL-C to exit"
			$runAgain = Read-Host -Prompt "Type [a] to search for updates again, press Enter to continue with the upgrade or press CTRL-C to exit"
			If ($runAgain -like "a"){
				continue
			}
			
			Write-Host "`nUpgrading packages" -ForegroundColor DarkRed

			$odTable | %{
				Write-Host "`n"
				.$printDate
				$_.Package
				$arguments = "upgrade " + $($_.Package + $(If($_.Parameters){" " + $_.Parameters})) + " -y --limit-output"
				# Debug:
				# $arguments
				Start-Process choco -ArgumentList $arguments -NoNewWindow -Wait
			}
			
			# Record currently installed packages for future reference
			Write-Host "`nWriting current installed packages to file..." -ForegroundColor DarkGray
			choco list --localonly --limitoutput | ConvertTo-Json | Out-File $logInstalled

			# Remove old log files
			Write-Host "Removing old log files..." -ForegroundColor DarkGray
			@($logDupes,$logInstalled,$logOutdated,$logTable) | %{
				$logs = $_ | %{
					Get-ChildItem -File $($_.Substring(0,($_.Length -20)) + '*') | Sort-Object -Property CreationTime -Descending | Select-Object -Skip 2
				}
				If ($logs.Count -gt 0) {$logs | Remove-Item -Force}
			}
			Read-Host -Prompt "Completed. Press Enter to exit"
		} Else {
			""
			Write-Host "Updates available but will not be performed due to exclusions" -ForegroundColor Yellow
			If ($dupes[$listDupes].Count -gt 0){
				""
				Write-Host "Duplicate packages being ignored:" -ForegroundColor Magenta
				$dupes[$listDupes] | Out-Host
			}
			If ($exclude.Count -gt 0){
				Write-Host "Excluded packages being ignored:" -ForegroundColor Magenta
				$exclude | Out-Host
			}
			
			# Remove old log files
			Write-Host "Removing old log files..." -ForegroundColor DarkGray
			@($logDupes,$logInstalled,$logOutdated,$logTable) | %{
				$logs = $_ | %{
					Get-ChildItem -File $($_.Substring(0,($_.Length -20)) + '*') | Sort-Object -Property CreationTime -Descending | Select-Object -Skip 2
				}
				If ($logs.Count -gt 0) {$logs | Remove-Item -Force}
			}
			
			# Ask to search again
			$runAgain = Read-Host -Prompt "Type [a] to search for updates again, press Enter to continue with the upgrade or press CTRL-C to exit"
			If ($runAgain -like "a"){
				continue
			}
		}
	} Else {
		Write-Host "No updates found" -ForegroundColor Green
		
		# Ask to search again
		$runAgain = Read-Host -Prompt "`nType [a] to search for updates again, press Enter or CTRL-C to exit"
	}
} While ($runAgain -like "a")