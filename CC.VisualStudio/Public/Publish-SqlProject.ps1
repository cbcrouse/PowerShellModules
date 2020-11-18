<#
.SYNOPSIS
	Publishes a SQL project in Visual Studio.
.DESCRIPTION
	This function finds a SQL project by name and attempts to build and publish it using the MSBuild executable.
.PARAMETER Name
	The name of the SQL project.
.PARAMETER PublishProfile
	The name of the publish xml file to use during the publish process.
.PARAMETER RemoveCachedFiles
	This option is the best attempt to resolve an issue with publishing the database
	due to a caching issue with the .dbmdl file (cached file of the project's db model).
	The project will be unloaded and reloaded if this option is enabled.
.EXAMPLE
	Publish-Project -Name "DB" -PublishProfile "staging.publish.xml" -RemoveCachedFiles
#>
Function Publish-SqlProject() {
	[CmdletBinding(
		SupportsShouldProcess,
		ConfirmImpact = 'High')]
	param(
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory = $true)]
		[string]$Name,
		[string]$PublishProfile = "Local.publish.xml",
		[switch]$RemoveCachedFiles
	)

	if ($PSCmdlet.ShouldProcess($MyInvocation.MyCommand)) {

		$script:Project = Get-SolutionProjects -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object { $_.Name -eq $Name } | Select-Object -First 1

		if ($null -eq $script:Project) {
			Write-Error "$($MyInvocation.MyCommand) stopped - project not found with name: '$Name'"
			return
		}

		if (-not $script:Project.FullName.EndsWith(".sqlproj")) {
			Write-Error "$($MyInvocation.MyCommand) stopped - invalid project path: '$($script:Project.FullName)'"
			return
		}

		$script:ProjectDirectory = Split-Path -Path $script:Project.FullName -Verbose:$VerbosePreference -ErrorAction Stop

		# Ensure MSBuild Tools are available
		$msBuildToolsDestination = "C:\Program Files (x86)\MSBuild\Microsoft\VisualStudio\v11.0\SSDT\"
		$msBuildToolsPath = "C:\Program Files (x86)\MSBuild\Microsoft\VisualStudio\v14.0\SSDT"

		if (-not $(Test-Path $msBuildToolsDestination)) {
			Copy-Item -Path $msBuildToolsPath -Destination $msBuildToolsDestination -Force -Recurse -Verbose:$VerbosePreference -ErrorAction 'Continue'
		}

		if ($RemoveCachedFiles) {

			$script:Project | Disable-Project -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference

			Remove-Item -Path "$script:ProjectDirectory\bin" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
			Remove-Item -Path "$script:ProjectDirectory\obj" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
			Remove-Item -Path "$script:ProjectDirectory\*.jfm" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
			Remove-Item -Path "$script:ProjectDirectory\*.dbmdl" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

			$UnloadedProject = Get-SolutionProjects -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
		
			if ($null -eq $UnloadedProject) {
				Write-Error "$($MyInvocation.MyCommand) stopped - could not find unloaded project with name: '$Name'"
				return
			}

			$UnloadedProject | Enable-Project -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
			$script:Project = Get-SolutionProjects -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
		}

		$FilePath = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
		$argumentList = "/t:ReBuild,Publish /p:SqlPublishProfilePath=$PublishProfile /p:SuppressTSqlWarnings=`"71502;71562;71558;`" `"$($script:Project.FullName)`""

		Invoke-Process -Path $FilePath -ArgumentList $argumentList -Verbose:$VerbosePreference -ErrorAction Stop
	}
}