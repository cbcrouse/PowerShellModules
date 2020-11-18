<#
.SYNOPSIS
	Get the currently open solution's name.
.DESCRIPTION
	Returns the name of the solution that is currently open.
.EXAMPLE
	C:\PS> Get-SolutionFullName
.NOTES
	Author:         Casey Crouse
	Created On:     07/22/2019
#>
Function Get-SolutionFullName() {
	[CmdletBinding()]
	param()

	if ($null -ne $dte) {
		if ($null -ne $dte.Solution) {
			$solutionName = $dte.Solution.Properties.Item("Name").Value
			$solutionPath = $dte.Solution.FullName
			Write-Verbose "Found solution: $solutionName"
			Write-Verbose "Returning: $solutionPath"
			return $solutionPath
		}
		else {
			Write-Warning "No solution is currently open."
			return $null
		}
	}

	Write-Error "Must be in Visual Studio to run this command."
	return $null
}