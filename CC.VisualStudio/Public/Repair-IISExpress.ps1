<#
.SYNOPSIS
	Remove localhost certificates and re-install IISExpress with missing files option.
.DESCRIPTION
	This function removes all localhost certificates and then 
	performs a re-installation of IISExpress with 'missing files' option selected in
	order to restore the localhost certificates.
.EXAMPLE
	Repair-IISExpress
.NOTES
	Author:         Casey Crouse
	Created On:     07/22/2019
#>
Function Repair-IISExpress() {
	[CmdletBinding(
		SupportsShouldProcess,
		ConfirmImpact = 'High')]
	param()

	Write-Warning `
		"Repair-IISExpress will break SSL bindings for other projects because 
		the localhost certificate will be removed. Open the other projects in 
		Visual Studio and use 'Reset-SSLBindings' to update the bindings with 
		the new localhost certificate.";

	if ($PSCmdlet.ShouldProcess($MyInvocation.MyCommand)) {
		Write-Verbose "Performing $($MyInvocation.MyCommand)"

		# Ensure that IISExpress is not open
		$IISExpressInstances = Get-Process -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object Name -Like iisexpress*
		if ($null -ne $IISExpressInstances) {
			Write-Error "$($MyInvocation.MyCommand) stopped - please close all instances of IIS Express to continue."
			return
		}

		# Get the localhost certificates
		$Certificates = Get-ChildItem -Path "Cert:\LocalMachine" -Recurse -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object Subject -EQ "CN=localhost"

		# Delete all localhost certificates
		if ($null -ne $Certificates) {
			foreach ($cert in $Certificates) {
				$output = Remove-Item -Path $cert.PSPath -Force -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Out-Null
				
				if ($output -eq -1) {
					Write-Warning "`tFailed to remove certificate: $($cert.GetName())"
				}
				else {
					Write-Verbose "`tCertificate was removed successfully."
				}
			}
		}

		Set-Location -Path C:\ -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
		
		Write-Verbose "Searching for application: IIS Express - this may take a minute..."
		$iisExpressObject = Get-WmiObject -Class Win32_Product -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object Name -like "IIS*Express"

		if ($null -ne $iisExpressObject) {

			Write-Verbose "Reinstalling application: $($iisExpressObject.Name)"

			try {
				$output = $iisExpressObject.Reinstall(1)

				if ($output.ReturnValue -eq 0) {
					Write-Verbose "`tRe-installation was successful!"
				}
				else {
					Write-Verbose "`tRe-installation was not successful! Return Code: $($output.ReturnValue). Check online here: https://msdn.microsoft.com/en-us/library/aa393044(v=vs.85).aspx for more information."
				}
			}
			catch [Exception]
			{
				Write-Error $_.Message
			}
		} else {
			Write-Verbose "IISExpress application not found."
		}
	}
}