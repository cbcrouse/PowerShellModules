<#
.SYNOPSIS
	Gets the virtual solution explorer path for the project.
.DESCRIPTION
	This function looks at the current object to determine it's virtual solution explorer path.
.PARAMETER Project
	The solution project object.
.EXAMPLE
	Get-SolutionExplorerProjectPath -Project [System.Object]
.EXAMPLE
	[System.Object] | Get-SolutionExplorerProjectPath
#>
Function Get-SolutionExplorerProjectPath() {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[System.Object]$Project
	)
	
	# Since this is a recursive function, there is no good way to avoid generating this 'gold' list each time.
	# However, this list is needed in order to figure out if a project has been unloaded by check its name against this list.
	# When a project is unloaded, sometimes, all we have is the name.
	Write-Verbose ""
	Write-Verbose "Building list of project files in order to get the project names 'gold' list."
	$script:slnFullName = Get-SolutionFullName -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
	$script:slnName = [System.IO.Path]::GetFileNameWithoutExtension($script:slnFullName)
	$script:ProjectNamesList = $slnFullName | Get-SolutionProjectNames -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference

	$script:partialPath = [string]::Empty

	$script:type = [Microsoft.VisualBasic.Information]::TypeName($Project)
	$script:typeObject = [Microsoft.VisualBasic.Information]::TypeName($Project.Object)
	$script:ProjectName = $Project.Name
	$script:ProjectFullName = $Project.FullName

	Write-Verbose "Project Name: $script:ProjectName"
	Write-Verbose "Project Full Name: $script:ProjectFullName"
	Write-Verbose "Project Object Type: $script:type"
	Write-Verbose "Project Object.Object Type: $script:typeObject"

	if ($script:typeObject -eq 'OAVSProject' -or $script:typeObject -eq 'OAProject' -or $script:typeObject -match '^VSProject(?<Version>\d*)$' -or
		$script:type -eq 'OAVSProject' -or $script:type -eq 'OAProject' -or	$script:type -match '^VSProject(?<Version>\d*)$')
	{
		Write-Verbose ""
		Write-Verbose "Project Type is in ('OAVSProject', 'OAProject', 'VSProject\d*') - checking to see if project has a parent..."
		Write-Verbose ""

		# Does the project have a parent? If so, it must have a containing project.
		if ($null -ne $Project.ParentProjectItem.ContainingProject) {
			$name = $Project.ParentProjectItem.ContainingProject.ProjectName
			$path = Join-Path -Path $name -ChildPath $Project.ParentProjectItem.Name
			Write-Verbose "Project.ParentProjectItem.ContainingProject found: $name"
			$script:partialPath = $path

		# If the project has no parent, but has the ContainingProject property, try this next.
		} elseif ($null -ne $Project.ContainingProject) {
			$name = $Project.ContainingProject.ProjectName
			$path = Join-Path -Path $name -ChildPath $Project.ContainingProject.ProjectName
			Write-Verbose "Project.ContainingProject found: $name"
			$script:partialPath = $path

		# The project must be at the root level, use the project's name.
		} else {
			Write-Verbose "No parent object found - project must be an item at the root level of solution."
			if ($null -ne $script:ProjectName) {
				Write-Verbose "Project found: $script:ProjectName"
				$script:partialPath = $script:ProjectName
			}
		}
	# This is most likely a project that is unloaded.
	} elseif ($script:type -eq 'ProjectItem' -and $script:typeObject -eq 'Nothing') {

		Write-Verbose ""
		Write-Verbose "Project Type is 'ProjectItem' and Project.Object Type is 'Nothing' - possible unloaded project."
		Write-Verbose ""

		$existingPath = Get-SolutionExplorerPathVariable -Name $script:ProjectName -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
		if (-not [string]::IsNullOrEmpty($existingPath)) {
			Write-Verbose "`tProject previously unloaded - using existing path."
			return $existingPath

		} elseif ($script:ProjectNamesList.Contains($script:ProjectName)) {
			$path = $script:ProjectName
			Write-Verbose "`tProject previously unloaded - no existing path found."
			Write-Verbose "`tIt is likely that this path is invalid unless it exists at the root of the solution."
			$script:partialPath = $path
		}
	# This is most likely a root solution project that is unloaded
	} elseif ($script:type -eq 'Project' -and $script:typeObject -eq 'Nothing') {

		Write-Verbose ""
		Write-Verbose "Project Type is 'Project' and Project.Object Type is 'Nothing' - possible unloaded project."
		Write-Verbose ""

		$existingPath = Get-SolutionExplorerPathVariable -Name $script:ProjectName -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
		if (-not [string]::IsNullOrEmpty($existingPath)) {
			Write-Verbose "`tProject previously unloaded - using existing path."
			return $existingPath

		} elseif ($script:ProjectNamesList.Contains($script:ProjectName)) {
			$path = $script:ProjectName
			Write-Verbose "`tProject previously unloaded - no existing path found."
			Write-Verbose "`tIt is likely that this path is invalid unless it exists at the root of the solution."
			$script:partialPath = $path
		}
	}

	$explorerPath = "$script:slnName\$script:partialPath"
	Write-Verbose "Returning explorer path: $explorerPath"
	return $explorerPath
}