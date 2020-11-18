<#
.SYNOPSIS
	Activates a window within Visual Studio UI.
.DESCRIPTION
	This function finds a window by its caption and activates it.
.PARAMETER Caption
	The caption used to search for the window. This value will likely be visible within the UI on the target window.
.EXAMPLE
	Set-FocusOnWindow -Caption "Solution Explorer"
.NOTES
	Author:         Casey Crouse
	Created On:     07/24/2019
#>
Function Set-FocusOnWindow() {
	[CmdletBinding()]
	param(
		[ValidateNotNullOrEmpty()]
		[Parameter(Mandatory = $true)]
		[string]$Caption
	)

	if ($null -eq $dte) {
		Write-Error "$($MyInvocation.MyCommand) stopped - this command requires Visaul Studio."
		return
	}
	
	Write-Warning "Performing UI Macro: Activating Window ($Caption)"

	$window = $dte.Windows | Where-Object { $_.Caption -eq $Caption }

	if ($null -eq $window) {
		Write-Error "$($MyInvocation.MyCommand) stopped - window with caption '$Caption' not found."
		return
	}

	$window.Activate()

	# WARNING: Do not put any code here that would change the focus of the UI to ensure the window stays active.
}