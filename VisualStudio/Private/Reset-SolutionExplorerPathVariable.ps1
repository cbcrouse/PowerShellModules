<#
.SYNOPSIS
	Finds and removes if exists, and sets a new global variable for a Solution Explorer path in Visual Studio.
.DESCRIPTION
	This function is specifically for holding on to Solution Explorer path values.
	After unloading a project, the path is no longer dynamically accessible and the
	only way to reload the project is to know the path.
.PARAMETER Name
	The name for the global variable. This should be the name of the project..
.PARAMETER Value
	The value for the global variable. This should be the path that is used to select the project in the Solution Explorer.
.EXAMPLE
	Reset-SolutionExplorerPathVariable -Name "MyFirst.WebApi" -Value "MySolutionName\src\Prentation\MyFirst.WebApi"
.NOTES
	Author:         Casey Crouse
	Created On:     07/23/2019
#>
Function Reset-SolutionExplorerPathVariable() {
	[CmdletBinding()]
	param(
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory=$true)]
		[string]$Name,
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory=$true)]
		[string]$Value
	)

	$pathName = "se_path_$Name"

	if (-not [string]::IsNullOrEmpty($Value)) {
		Write-Verbose "Storing the following as a global variable:"
		Write-Verbose "`tName: $pathName"
		Write-Verbose "`tValue: $Value"
		
		$variable = $(Get-Variable -Name $pathName -Scope Global -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Out-Null).Value
		if (-not [string]::IsNullOrEmpty($variable)) {
			Remove-Variable $pathName -Scope Global -Force -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Out-Null
			Set-Variable -Name $pathName -Value $Value -Option ReadOnly -Scope Global -Visibility Public -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Out-Null
		} else {
			Set-Variable -Name $pathName -Value $Value -Option ReadOnly -Scope Global -Visibility Public -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Out-Null
		}
	}
}