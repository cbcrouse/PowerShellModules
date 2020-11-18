<#
.SYNOPSIS
	Close the current solution
.DESCRIPTION
	Closes the currently open solution.
.EXAMPLE
	C:\PS> Close-Solution
.NOTES
	Author:         Casey Crouse
	Created On:     10/18/2019
#>
Function Close-Solution {
	[CmdletBinding()]
	param()

	if ($null -ne $dte) {
		if (-not $([string]::IsNullOrEmpty($dte.Solution.FullName))) {
			$solutionName = $dte.Solution.FullName
			Write-Warning "Closing $solutionName"
			$dte.Solution.Close($true);
			return
		}
		else {
			return
		}
	}
	
	throw "Must be in Visual Studio to run this command."
}