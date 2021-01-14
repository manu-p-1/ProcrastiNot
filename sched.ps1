<#
Filename: sched.ps1
Author: Manu P.

This file runs the ProcrastiNot script as mentioned in the README.md file
#>

param (
	[Parameter(Mandatory=$true)][string]$Filename,
	[Parameter(Mandatory=$true)][string]$ZoomExe
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

# Check the Filename
if (-not $Filename.endswith(".pn") )
{
	ShowError "The Filename was not recognized. A .pn file must be provided" $true
}

# Open the file and read contents
try{
	$lines = [System.IO.File]::ReadLines($Filename)
} catch [System.Management.Automation.MethodInvocationException], [System.IO.IOException]{
	ShowError "The .pn file provided could not be found or opened" $true
}

# Load the names into a hashtable and check for any duplicates
$dowAbbrevToFull = @{
	Mon = [System.DayOfWeek]::Monday;
	Tue = [System.DayOfWeek]::Tuesday;
	Wed = [System.DayOfWeek]::Wednesday;
	Thu = [System.DayOfWeek]::Thursday;
	Fri = [System.DayOfWeek]::Friday;
	Sat = [System.DayOfWeek]::Saturday;
	Sun = [System.DayOfWeek]::Sunday
}

$names = @{}
$ctr = 0
foreach($line in $lines)
{
	$ctr++
	if(-not $line){ 
		continue
	}
	$components = $line -Split ","
	if(-not $components.Length -eq 4){
		ShowError "The .pn file is improperly formatted" $true
	}
	if($names.ContainsKey($components[0])){
		ShowError "The .pn file contains a duplicate name label on line ${ctr}. The process will continue and ignore this line."
		continue
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
		ShowError "The .pn file contains a malformed date ${ctr}. The process will continue and ignore this line."
		continue
	}

	# Schedule Day of Week
	$days = $cmp_day.Split("/")
	$days = $days | Select-Object -Unique | ForEach-Object {
		$trimmed = $_.Trim()
		$tx = (Get-Culture).TextInfo.ToTitleCase($trimmed)
		$dowAbbrevToFull[$tx]
	} 

	foreach($day in $days){
		# Schedule Name
		$schedName = "ProcrastiNot_TASK__$($cmp_name)_$($day)"

		# Schedule link
		$schedLink = $cmp_link

		# Task Action
		$taskAction = New-ScheduledTaskAction -Execute "$($ZoomExe)" -Argument "-url=$($schedLink)"
		
		# Task Trigger
		$taskTrigger = New-ScheduledTaskTrigger `
						-Weekly `
						-DaysOfWeek $days `
						-At $startTime 
		
		# Task Desc.
		$description = "ProcrastiNot Zoom Starter Information: Date:$($startTime.ToString('h:mm tt')), Zoom:$($schedLink)"

		# Task Existence
		$taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -eq $schedName }
		
		if(-Not $taskExists) {
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
}

<# 
To Unregister, type this into PowerShell:
Get-ScheduledTask "ProcrastiNot_TASK__*" | Unregister-ScheduledTask
#>