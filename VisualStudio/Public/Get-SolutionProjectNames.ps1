<#
.SYNOPSIS
	Returns a list of project names for a given solution.
.DESCRIPTION
	This function reads the solution file and returns a list of names for all projects.
.PARAMETER Path
	The path to the solution file.
.EXAMPLE
	Get-SolutionProjectNames -Path "C:\Git\MySolution\MySolution.sln"
.EXAMPLE
	"C:\Git\MySolution\MySolution.sln" | Get-SolutionProjectNames
.NOTES
	Author:         Casey Crouse
	Created On:     07/24/2019
#>
Function Get-SolutionProjectNames() {
	[CmdletBinding()]
	param(
		[ValidateScript({ $_.EndsWith(".sln")})]
		[ValidateScript({ Test-Path -Path $_ -Type Leaf })]
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[string]$Path
	)

	$SlnFileContent = Get-Content -Path $Path -Force -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference

	if ($null -eq $SlnFileContent) {
		return $null
	}

	$RegexPattern = '^Project\("(?<ProjectTypeId>[a-zA-Z0-9}{_-]*)"\)\s=\s"(?<ProjectName>.*)",\s"(?<RelativePath>.*)",\s"{(?<ProjectId>[a-zA-Z0-9}{_-]*)}"$'

	$ProjectNamesList = New-Object System.Collections.Generic.List[string] -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference

	foreach($line in $SlnFileContent) {
		$selections = $line | Select-String -Pattern $RegexPattern
		if ($null -ne $selections.Matches) {
			foreach($match in $selections.Matches) {
				$matchName = $($match.Groups["ProjectName"]).ToString().Trim()
				$matchPath = $($match.Groups["RelativePath"]).ToString().Trim()
				if ($matchName -ne $matchPath) {
					Write-Verbose "Match: { Name: $matchName, Path: $matchPath }"
					$ProjectNamesList.Add($matchName)
				} else {
					Write-Verbose "`Skip (Non-Project): $matchName"
				}
			}
		}
	}

	return $ProjectNamesList
}