<#
.SYNOPSIS
	Restart the current solution
.DESCRIPTION
	Closes the currently open solution and reopens it.
.PARAMETER Rebuild
	Clean, restore, and build the solution using the dotnet CLI.
.EXAMPLE
	C:\PS> Restart-Solution
.NOTES
	Author:         Casey Crouse
	Created On:     03/06/2019
#>
Function Restart-Solution {
	[CmdletBinding()]
	param(
		[switch]$Rebuild
	)

	if ($null -ne $dte) {
		if (-not $([string]::IsNullOrEmpty($dte.Solution.FullName))) {
			$solutionName = $dte.Solution.FullName

			if ($Rebuild) {
				dotnet clean "$solutionName"; dotnet restore "$solutionName"; dotnet build "$solutionName";
			}

			Write-Host "Closing $solutionName..." -ForegroundColor DarkYellow
			$dte.Solution.Close($true);
			Write-Host "Opening $solutionName..." -ForegroundColor DarkGreen
			$dte.Solution.Open($solutionName)
			return
		}
		else {
			Write-Host "No solution currently open to restart." -ForegroundColor DarkRed
			return
		}
	}

	Write-Host "Must be in Visual Studio to run this command." -ForegroundColor DarkRed
	return
}