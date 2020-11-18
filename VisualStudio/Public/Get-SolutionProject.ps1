<#
.SYNOPSIS
	Get a solution project object.
.DESCRIPTION
	Returns the object that represents the Visual Studio project.
.PARAMETER Name
	The name of the project.
.PARAMETER ExactMatch
	Determines whether or not to find the project by exact match.
.EXAMPLE
	The following example will search for "Web" using several search conditions.

	C:\PS> Get-SolutionProject -Name "Web"
.EXAMPLE
	The following example will search for "Web" exactly as case-insensitive.

	C:\PS> Get-SolutionProject -Name "Web" -ExactMatch
.NOTES
	Author:         Casey Crouse
	Created On:     07/18/2019
#>
Function Get-SolutionProject {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Name,
		[switch]$ExactMatch
	)

	if ($null -eq $dte){
		Write-Error "$($MyInvocation.MyCommand) stopped - this command requires Visaul Studio."
		return $null
	}

	# Standardizing checks against lowercase only
	$script:Name = $Name.ToLower()

	$projects = Get-SolutionProjects -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference

	Write-Verbose "Searching visual studio projects for: $script:Name"
	Write-Verbose " "

	if ($ExactMatch) {
		foreach($project in $projects) {
			Write-Verbose "Name: $($project.Name)"
			$local:projectName = $project.Name.ToLower()

			if ($null -ne $project -and $local:projectName -eq $script:Name) {
				Write-Verbose "Found project exact match: $($project.Name)"
				return $project;
			}
		}
	} else {
		# Search for the project name using other search conditions.
		foreach($project in $projects) {
			Write-Verbose "Name: $($project.Name)"
			$local:projectName = $project.Name.ToLower()
			# Project names can sometimes contain several prefixes such as
			# CompanyName.Prefix2.Prefix3.Web where 'Web' is the name that is searched for.
			$lastNameSection = $local:projectName.Split(".")[-1].ToLower()

			Write-Verbose "Compare Items:"
			Write-Verbose "`tLast Section: $lastNameSection"
			Write-Verbose "`tName: $local:projectName"
			Write-Verbose " "

			if ($null -ne $project -and (
				$lastNameSection -eq $script:Name -or <# Name Match #>
				$local:projectName -eq $script:Name <# Last Section Match #>
				<# Add new search conditions here #>
			)) {
				Write-Verbose "Found project match: $($project.Name)"
				return $project;
			}
		}
	}

	Write-Verbose "Project not found."
	return $null;
}