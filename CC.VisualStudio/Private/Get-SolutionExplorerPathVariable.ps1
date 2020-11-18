
<#
.SYNOPSIS
	Finds and removes if exists, and sets a new global variable for a Solution Explorer path in Visual Studio.
.DESCRIPTION
	This function is specifically for retrieving Solution Explorer path values.
	After unloading a project, the path is no longer dynamically accessible and
	the only way to reload the project is to know the path.
.PARAMETER Name
	The name of the global variable.
.EXAMPLE
	Get-SolutionExplorerPathVariable -Name "MyFirst.WebApi"
.NOTES
	Author:         Casey Crouse
	Created On:     07/23/2019
#>
Function Get-SolutionExplorerPathVariable() {
	[CmdletBinding()]
	param(
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory=$true)]
		[string]$Name
	)

	$pathName = "se_path_$Name"

	return Get-Variable -Name $pathName -Scope Global -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Select-Object -ExpandProperty Value
}