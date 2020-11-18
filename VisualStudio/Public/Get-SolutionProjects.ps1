<#
.SYNOPSIS
	Recursively retrieve solution projects.
.DESCRIPTION
	This function navigates through the solution $dte object finding projects and recursively returning them.
.PARAMETER Object
	Represents a System.__ComObject from the $dte.
	The Object does not need to be set for this function to succeed.
	If this value is not set, this function will begin using $dte.Solution.Projects.
.EXAMPLE
	Get-SolutionProjects
.EXAMPLE
	Get-SolutionProjects -Verbose
.NOTES
	Author:         Casey Crouse
	Created On:     07/22/2019
#>
Function Get-SolutionProjects() {
	[CmdletBinding()]
	param(
		[System.Object]$Object
	)

	if ($null -eq $dte) {
		Write-Error "$($MyInvocation.MyCommand) stopped - this command requires Visaul Studio."
		return $null
	}

	if ($null -eq $Object) {
		# Nothing was passed in, this is likely to be the first iteration.
		$Object = $dte.Solution.Projects

		# Create the new list for projects now
		$script:projectList = New-Object System.Collections.Generic.List[System.Object]
		Write-Verbose "Creating new list for the first time."
		Write-Verbose ""
	
		# The following ProjectNamesList is needed in order to figure out if a project has been unloaded.
		# When a project is unloaded, sometimes, all we have is the name.
		# We can check the name, if we have it, against the 'gold' list to ensure it's an actual project and not a virtual folder.
		Write-Verbose ""
		Write-Verbose "Building list of project files in order to get the project names 'gold' list."
		# "Script" scoped values will keep their values through the recursion.
		$script:SolutionDirectory = Split-Path -Path $(Get-SolutionFullName -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference)
		$script:ProjectNamesList = New-Object System.Collections.Generic.List[string] -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
		$script:ProjectFileNames = Get-ChildItem -Path $script:SolutionDirectory -Recurse -Force -Include "*.*proj" -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
	
		Write-Verbose "Adding the following names:"
		foreach ($name in $script:ProjectFileNames) {
			$withoutExt = [System.IO.Path]::GetFileNameWithoutExtension($name)
			Write-Verbose "`t$name"
			$script:ProjectNamesList.Add($withoutExt)
		}
		Write-Verbose ""
	}

	Write-Verbose "Project list count is: $($script:projectList.Count)"
	Write-Verbose ""
	
	# Get the entry object types
	$script:type = [Microsoft.VisualBasic.Information]::TypeName($Object)
	$script:typeObject = [Microsoft.VisualBasic.Information]::TypeName($Object.Object)
	$script:ProjectName = $Object.Name
	$script:ProjectFullName = $Object.FullName

	Write-Verbose "Project Name: $script:ProjectName"
	Write-Verbose "Project Full Name: $script:ProjectFullName"
	Write-Verbose "Project Object Type: $script:type"
	Write-Verbose "Project Object.Object Type: $script:typeObject"
	
	# Is the item coming in, a project? Check it's type
	if ($script:type -eq 'OAVSProject' -or $script:type -eq 'OAProject' -or $script:type -eq 'VSProject3') {
		Write-Verbose ""
		Write-Verbose "`tAdding to list: $($Object.Name)"
		$script:projectList.Add([System.Object]$Object)
		Write-Verbose ""
	# Is the item coming in, a project? Check it's object's type
	} elseif ($script:typeObject -eq 'OAVSProject' -or $script:typeObject -eq 'OAProject' -or $script:typeObject -eq 'VSProject3') {
		Write-Verbose ""
		Write-Verbose "`tAdding to list: $($Object.Name)"
		$script:projectList.Add([System.Object]$Object.Object)
		Write-Verbose ""

	# Projects is the root collection
	} elseif ($script:type -eq 'Projects') {
		Write-Verbose ""
		Write-Verbose "Entered IF BLOCK 'Projects'..."

		foreach ($project in $Object) {
			$local:type = [Microsoft.VisualBasic.Information]::TypeName($project)
			$local:typeObject = [Microsoft.VisualBasic.Information]::TypeName($project.Object)

			Write-Verbose ""
			Write-Verbose "`tWithin Projects - Name: $($project.Name)"
			Write-Verbose "`t`tEntry Object Type: $local:type"
			Write-Verbose "`t`tEntry Object.Object Type: $local:typeObject"

			# We're at the top level folders, each folder below is considered a 'project'
			if ($local:typeObject -eq 'SolutionFolder') {
				Write-Verbose "`t`tType is 'SolutionFolder' - Starting Recursion on $($project.Name)"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $project.ProjectItems | Out-Null

			# We have a project at the top level of the solution
			} elseif ($local:type -eq 'OAVSProject' -or $local:type -eq 'OAProject' -or $local:type -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($project.Name)"
				$script:projectList.Add([System.Object]$project)
				Write-Verbose ""
			} elseif ($local:typeObject -eq 'OAVSProject' -or $local:typeObject -eq 'OAProject' -or $local:typeObject -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($project.Name)"
				$script:projectList.Add([System.Object]$project)
				Write-Verbose ""
			} elseif ($local:type -eq 'Project' -and $local:typeObject -eq 'Nothing') {
				if ($script:ProjectNamesList.Contains($project.Name)) {
					# Found a project type that might be unloaded. Unable to check "UniqueName" here
					# because only some of the project types have this property. Name, however is available.
					# Consider checking the name in a list of project files that we can gather at the start.
					Write-Verbose ""
					Write-Verbose "`t`t$($project.Name) is unloaded."
					Write-Verbose "`t`tAdding to list: $($project.Name)"
					$script:projectList.Add([System.Object]$project)
					Write-Verbose ""
				}
			} elseif ($local:type -eq 'Project' -and $local:typeObject -ne 'Nothing') {
				Write-Verbose "`tType is 'Project' - Starting Recursion on $($project.Name)"
				Write-Verbose "`t$([Microsoft.VisualBasic.Information]::TypeName($project.ProjectItems))"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $project.ProjectItems | Out-Null

			}
			else {
				Write-Information "$($project.Name) unknown type: $local:typeObject"
			}
		}
	} elseif ($script:type -eq 'ProjectItems') {
		Write-Verbose ""
		Write-Verbose "Entered IF BLOCK 'ProjectItems'..."

		foreach ($project in $Object) {

			$local:type = [Microsoft.VisualBasic.Information]::TypeName($project)
			$local:typeObject = [Microsoft.VisualBasic.Information]::TypeName($project.Object)

			Write-Verbose ""
			Write-Verbose "`tWithin ProjectItems - Name: $($project.Name)"
			Write-Verbose "`t`tEntry Object Type: $local:type"
			Write-Verbose "`t`tEntry Object.Object Type: $local:typeObject"
			
			if ($local:type -eq 'ProjectItem' -and $local:typeObject -eq 'Nothing') {
				if ($script:ProjectNamesList.Contains($project.Name)) {
					# Found a project type that might be unloaded. Unable to check "UniqueName" here
					# because only some of the project types have this property. Name, however is available.
					# Consider checking the name in a list of project files that we can gather at the start.
					Write-Verbose ""
					Write-Verbose "`t`t$($project.Name) is unloaded."
					Write-Verbose "`t`tAdding to list: $($project.Name)"
					$script:projectList.Add([System.Object]$project)
					Write-Verbose ""
				}
			} elseif ($local:type -eq 'Project') {
				Write-Verbose "`tType is 'Project' - Starting Recursion on $($project.Name)"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $project | Out-Null

			} elseif ($local:type -eq 'OAVSProject' -or $local:type -eq 'OAProject' -or $local:type -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($project.Name)"
				$script:projectList.Add([System.Object]$project)
				Write-Verbose ""
			} elseif ($local:typeObject -eq 'OAVSProject' -or $local:typeObject -eq 'OAProject' -or $local:typeObject -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($project.Name)"
				$script:projectList.Add([System.Object]$project.Object)
				Write-Verbose ""
			} elseif ($local:type -eq 'ProjectItem' -and $local:typeObject -ne 'Nothing') {
				Write-Verbose "`tType is 'ProjectItem' - Starting Recursion on $($project.Name)"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $project | Out-Null

			} else {
				Write-Information "$($project.Name) unknown type: $local:type"
			}
		}
	} elseif ($script:type -eq 'ProjectItem') {

		Write-Verbose ""
		Write-Verbose "Entered IF BLOCK 'ProjectItem'..."
		
		foreach ($projectItem in $Object.Object.ProjectItems) {

			$local:type = [Microsoft.VisualBasic.Information]::TypeName($projectItem)
			$local:typeObject = [Microsoft.VisualBasic.Information]::TypeName($projectItem.Object)

			Write-Verbose ""
			Write-Verbose "`tWithin ProjectItem - Name: $($projectItem.Name)"
			Write-Verbose "`t`tEntry Object Type: $local:type"
			Write-Verbose "`t`tEntry Object.Object Type: $local:typeObject"
			
			if ($local:type -eq 'ProjectItem' -and $local:typeObject -eq 'Nothing') {
				if ($script:ProjectNamesList.Contains($projectItem.Name)) {
					# Found a project type that might be unloaded. Unable to check "UniqueName" here
					# because only some of the project types have this property. Name, however is available.
					# Consider checking the name in a list of project files that we can gather at the start.
					Write-Verbose ""
					Write-Verbose "`t`t$($projectItem.Name) is unloaded."
					Write-Verbose "`t`tAdding to list: $($projectItem.Name)"
					$script:projectList.Add([System.Object]$projectItem)
					Write-Verbose ""
				}
			} elseif ($local:typeObject -eq 'Project') {
				Write-Verbose "`t`tType is 'Project' - Starting Recursion on $($projectItem.Name)"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $projectItem | Out-Null

			} elseif ($local:type -eq 'OAVSProject' -or $local:type -eq 'OAProject' -or $local:type -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($projectItem.Name)"
				$script:projectList.Add([System.Object]$projectItem)
				Write-Verbose ""
			} elseif ($local:typeObject -eq 'OAVSProject' -or $local:typeObject -eq 'OAProject' -or $local:typeObject -match '^VSProject(?<Version>\d*)$') {
				Write-Verbose ""
				Write-Verbose "`tAdding to list: $($projectItem.Name)"
				$script:projectList.Add([System.Object]$projectItem.Object)
				Write-Verbose ""
			} elseif ($local:typeObject -eq 'ProjectItem') {
				Write-Verbose "`tType is 'ProjectItem' - Starting Recursion on $($projectItem.Name)"

				# Recursion - Capturing return list to avoid bad output.
				# However, since the projectList variable scope is 'script',
				# the items have already been added to the list, no need to re-add them.
				$items = Get-SolutionProjects -Object $projectItem | Out-Null

			} else {
				Write-Information "$($Object.Name) unknown type: $local:typeObject"
			}
		}
	}
	
	Write-Verbose "List new count is...$($script:projectList.Count)"

	return $script:projectList
}