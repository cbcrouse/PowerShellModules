<#
.SYNOPSIS
	Unloads a project within Visual Studio.
.DESCRIPTION
	This function selects a project in the solution explorer and executes the 'Project.UnloadProject' command.
.PARAMETER Project
	The solution project object.
.EXAMPLE
	Disable-Project -Project [System.Object]
.EXAMPLE
	[System.Object] | Disable-Project
.NOTES
	Author:         Casey Crouse
	Created On:     07/23/2019
#>
Function Disable-Project() {
	[CmdletBinding()]
	param(
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[System.Object]$Project
	)

	$script:ProjectName = $Project.Name

	Write-Verbose "Starting to disable project: $script:ProjectName"
	$script:explorerPath = $($Project | Get-SolutionExplorerProjectPath -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference)

	if ([string]::IsNullOrEmpty($script:explorerPath)) {
		Write-Warning "$($MyInvocation.MyCommand) stopped - virtual solution explorer path not found."
		return
	}

	Write-Verbose "Solution Explorer Path: $script:explorerPath"
	Write-Warning "Performing UI Macro: DO NOT CLICK ANYWHERE or the Solution Explorer may lose focus!"

	# VERY IMPORTANT: DO NOT PUT ANY CODE AFTER THIS POINT THAT WILL CHANGE THE UI FOCUS
	# EXAMPLE: Write-Verbose will focus the console.

	Set-FocusOnWindow -Caption "Solution Explorer"
	# These next actions must be as close as possible to each other.
	$dte.ActiveWindow.Object.GetItem($script:explorerPath).Select(1);
	$dte.ExecuteCommand("Project.UnloadProject")
	Set-FocusOnWindow -Caption "Package Manager Console"

	# Setting a global variable to keep track of the explorer virtual path after disabling the project.
	Reset-SolutionExplorerPathVariable -Name $script:ProjectName -Value $script:explorerPath -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
}