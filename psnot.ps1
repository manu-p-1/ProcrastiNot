<#
Filename: psnot.ps1
Author: Manu P.

This file runs the ProcrastiNot script as mentioned in the README.md file.
To Unregister all ProcrastiNot tasks, type the following piece of code into your PowerShell:
Get-ScheduledTask "ProcrastiNot_TASK__*" | Unregister-ScheduledTask
#>

param (
	[Parameter()][string]$Filename=(Join-Path $PSScriptRoot ".\sched.txt"),
	[Parameter()][string]$ZoomExe=([System.IO.Path]::GetFullPath([Environment]::GetFolderPath('ApplicationData') + "\Zoom\bin\Zoom.exe"))
)

<#
Shows an error prompt then promptly exits the program
#>
function ShowError([string]$msg, [bool]$exit_prog=$false){
	[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
	[System.Windows.Forms.MessageBox]::Show($msg, 
	"ProcrastiNot Error",
	[System.Windows.Forms.MessageBoxButtons]::OK,
	[System.Windows.Forms.MessageBoxIcon]::Error)
	
	if($exit_prog){
		exit 1
	}
}

# Check .pn in case and provide friendly warning
if ($Filename.EndsWith(".pn")){
	ShowError "Support for .pn files ended with version 1.0.0. Download a more recent version" $true
}

# Check the Filename
if (-not $Filename.endswith(".txt") )
{
	ShowError "The Filename was not recognized. A .txt file must be provided" $true
}

# Open the file and read contents
try{
	$lines = [System.IO.File]::ReadLines($Filename)
} catch [System.Management.Automation.MethodInvocationException], [System.IO.IOException]{
	ShowError "The .txt file provided could not be found or opened" $true
}

# Load the abbrevations into a hashtable
$dowAbbrevToFull = @{
	Mon = [System.DayOfWeek]::Monday;
	Tue = [System.DayOfWeek]::Tuesday;
	Wed = [System.DayOfWeek]::Wednesday;
	Thu = [System.DayOfWeek]::Thursday;
	Fri = [System.DayOfWeek]::Friday;
	Sat = [System.DayOfWeek]::Saturday;
	Sun = [System.DayOfWeek]::Sunday
}

$names = [System.Collections.ArrayList]@()
$ctr = 0
foreach($line in $lines)
{
	$ctr++
	if(-not $line){ 
		continue
	}

	# Load the names into a hashtable and check for any duplicates
	$components = $line -split ","
	if(-not $components.Length -eq 4){
		ShowError "The .txt file is improperly formatted" $true
	}
	if($names -contains $components[0]){
		ShowError "The .txt file contains a duplicate name label on line ${ctr}. The process will continue and ignore this line."
		continue
	} else {
		[void]$names.Add($components[0])
	}

	# Component names
	$cmp_name = $components[0].Trim()
	$cmp_time = $components[1].Trim()
	$cmp_day = $components[2]
	$cmp_link = $components[3].Trim()

	# Schedule Start Time
	try{
		$startTime = Get-Date -Date $cmp_time
	} catch {
		ShowError "The .txt file contains a malformed date ${ctr}. The process will continue and ignore this line."
		continue
	}

	# Schedule Day of Week
	$cmp_day = $cmp_day -replace '\s',''
	if($cmp_day -match "^([*])$"){
		$cmp_day = $cmp_day.Replace("*", "Mon/Tue/Wed/Thu/Fri")
	}
	if($cmp_day -match "^([*][*])$"){
		$cmp_day = $cmp_day.Replace("**", "Mon/Tue/Wed/Thu/Fri/Sat/Sun")
	}
	if($cmp_day -match "^([$])$"){
		$cmp_day = $cmp_day.Replace("$", "Mon/Wed/Fri")
	}
	if($cmp_day -match "^([$][$])$"){
		$cmp_day = $cmp_day.Replace("$$", "Tue/Thu")
	}
	if($cmp_day -match "^([!])$"){
		$cmp_day = $cmp_day.Replace("ends", "Sat/Sun")
	}

	$cmp_day = $cmp_day.Replace("//", "/")
	$days = $cmp_day -split "/"

	$days = $days | Select-Object -Unique | ForEach-Object {
		$trimmed = $_.Trim()
		$tx = (Get-Culture).TextInfo.ToTitleCase($trimmed)
		try{
			$dowAbbrevToFull[$tx]
		} catch{
			ShowError "The .txt file contains a day of week. The process will continue and ignore this line."
			continue
		}
	} 

	foreach($day in $days){
		# Schedule Name
		$schedName = "ProcrastiNot_TASK__$($cmp_name)_$($day)"

		# Task Precheck
		$taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -eq $schedName }
		if($taskExists) {
			Write-Host "The Task: $($schedName) already exists" 
			continue
		}

		# Schedule Link
		$schedLink = $cmp_link

		# Task Action
		$taskAction = New-ScheduledTaskAction -Execute "$($ZoomExe)" -Argument "-url=$($schedLink)"
		
		# Task Trigger
		$taskTrigger = New-ScheduledTaskTrigger `
						-Weekly `
						-DaysOfWeek $day `
						-At $startTime 
		
		# Task Desc.
		$description = "ProcrastiNot Zoom Starter Information --- Date: $($startTime.ToString('h:mm tt')), Zoom: $($schedLink)"

		# Task Register
		try{
			Register-ScheduledTask `
				-TaskName $schedName `
				-Action $taskAction `
				-Trigger $taskTrigger `
				-Description $description
			}
		catch{
			ShowError "The task could not be created. Check your privelages" $true
		}
	}
}
