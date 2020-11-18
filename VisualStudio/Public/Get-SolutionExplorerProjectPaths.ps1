<#
.SYNOPSIS
	Recursively retrieve solution explorer project paths.
.DESCRIPTION
	This function navigates through the solution $dte object finding projects and recursively
	returns paths to the projects. This is useful when trying to select projects in the
	solution explorer window.
.EXAMPLE
	Get-SolutionExplorerProjectPaths
.EXAMPLE
	Get-SolutionExplorerProjectPaths -Verbose
#>
Function Get-SolutionExplorerProjectPaths() {
	[CmdletBinding()]
	param(
	)

	$projects = Get-SolutionProjects -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
	$script:list = New-Object System.Collections.Generic.List[string]

	foreach($project in $projects) {
		$path = $($project | Get-SolutionExplorerProjectPath -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference)
		if (-not [string]::IsNullOrEmpty($path)) {
			$script:list.Add($path)
		}
	}

	return $script:list
}