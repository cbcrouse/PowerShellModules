<#
.SYNOPSIS
    Reset SSL bindings of ports from certificates.
.DESCRIPTION
	This function finds port bindings for specificied certificates, removes them, and adds them back.
.PARAMETER Ports
    The ports to reset bindings for.
.PARAMETER CertificateName
    The name of the certificate to reset SSL bindings.
.PARAMETER StoreName
    The store name for the certificate.
.EXAMPLE
	Reset-SSLBindings -PortBase "5000"
.EXAMPLE
    Reset-SSLBindings -PortBase "5000" -SecurePort "5001"
.NOTES
    Author:         Casey Crouse
    Created On:     07/22/2019

netsh syntax and parameter info:

Can only choose one port: hostnameport=<name:port> | ipport=<ipaddr:port> | ccs=<port> appid=<GUID>

Usage: add sslcert hostnameport=<name:port> | ipport=<ipaddr:port> | ccs=<port> appid=<GUID>
            [certhash=<string>] [certstorename=<string>] [verifyclientcertrevocation=enable|disable]
            [verifyrevocationwithcachedclientcertonly=enable|disable] [usagecheck=enable|disable] [revocationfreshnesstime=<u-int>]
            [urlretrievaltimeout=<u-int>] [sslctlidentifier=<string>] [sslctlstorename=<string>] [dsmapperusage=enable|disable]
            [clientcertnegotiation=enable|disable] [reject=enable|disable] [disablehttp2=enable|disable] [disablequic=enable|disable]
            [disablelegacytls=enable|disable] [disabletls12=enable|disable] [disabletls13=enable|disable] [disableocspstapling=enable|disable]
            [enabletokenbinding=enable|disable] [logextendedevents=enable|disable] [enablesessionticket=enable|disable]

Parameters: Tag: Value
            ipport: IP address and port for the binding.
            hostnameport: Unicode hostname and port for binding.
            ccs: Central Certificate Store binding.
            certhash: The SHA hash of the certificate. This hash is 20 bytes long and specified as a hex string.
            appid: GUID to identify the owning application.
            certstorename: Store name for the certificate. Required for Hostname based configurations. Defaults to MY for IP based configurations. Certificate must be stored in the local machine context.
            verifyclientcertrevocation: Turns on/off verification of revocation of client certificates.
            verifyrevocationwithcachedclientcertonly: Turns on/off usage of only cached client certificate for revocation checking.
            usagecheck: Turns on/off usage check. Default is enabled.
            revocationfreshnesstime: Time interval to check for an updated certificate revocation list (CRL). If this value is 0, then the new CRL is updated only if the previous one expires (in seconds).
            urlretrievaltimeout: Timeout on attempt to retrieve certificate revocation list for the remote URL (in milliseconds).
            sslctlidentifier: List the certificate issuers that can be trusted. This list can be a subset of the certificate issuers that are trusted by the machine.
            sslctlstorename: Store name under LOCAL_MACHINE where SslCtlIdentifier is stored.
            dsmapperusage: Turns on/off DS mappers. Default is disabled.
            clientcertnegotiation: Turns on/off negotiation of certificate. Default is disabled.
            reject: When enabled, any new matching connection is immediately dropped.
            disablehttp2: When set, HTTP2 is disabled for new matching connections immediately.
            disablequic: When set, QUIC is disabled for new matching connections immediately.
            disablelegacytls: When set, legacy versions of TLS are disabled.
            disabletls12: When set, TLS1.2 is disabled for new matching connections immediately.
            disabletls13: When set, TLS1.3 is disabled for new matching connections immediately.
            disableocspstapling: When set, OCSP stapling is disabled for new matching connections immediately.
            enabletokenbinding: When set, token binding is enabled for new connections immediately.
            logextendedevents: When set, additional events useful for debugging are logged.
            enablesessionticket: When set, TLS session resumption is enabled. 
                
            Remarks: Adds an SSL server certificate binding and corresponding client certificate policies for an IP address or hostname and a port.
            Examples: add sslcert ipport=1.1.1.1:443 certhash=0102030405060708090A appid={00112233-4455-6677-8899-AABBCCDDEEFF}
                      add sslcert hostnameport=www.contoso.com:443 certhash=0102030405060708090A appid={00112233-4455-6677-8899-AABBCCDDEEFF} certstorename=MY
                      add sslcert scopedccs=www.contoso.com:443 appid={00112233-4455-6677-8899-AABBCCDDEEFF}
                      add sslcert ccs=443 appid={00112233-4455-6677-8899-AABBCCDDEEFF}

#>
function Reset-SSLBindings() {
    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High')]
    param(
        [ValidatePattern("^([0-9]{1,4}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])$")]
        [ValidateScript({ $_.Count -gt 0 })]
        [Parameter(Mandatory = $true)]
        [int[]]$Ports,
        [ValidateNotNullOrEmpty()]
        $CertificateName = "CN=localhost",
        [ValidateNotNullOrEmpty()]
        $StoreName = "My"
    )

	if ($PSCmdlet.ShouldProcess($MyInvocation.MyCommand)) {
		Write-Verbose "Performing $($MyInvocation.MyCommand)"

        # Certificate must be stored in the local machine context.
        $LocalHostCerts = $(Get-ChildItem -Path "Cert:\LocalMachine\$StoreName" -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference | Where-Object Subject -eq $CertificateName)
        
        foreach ($localhostCert in $LocalHostCerts) {

            Write-Verbose "Certificate Details:"
            Write-Verbose "`tFriendly Name: $($localhostCert.FriendlyName)"
            Write-Verbose "`tName: $($localhostCert.GetName())"
            Write-Verbose "`tPSPath: $($localhostCert.PSPath)"
            Write-Verbose "`tThumprint: $($localhostCert.Thumbprint)"
            Write-Verbose " "

            foreach ($portNumber in $Ports) {

                $local:guid = [guid]::NewGuid().ToString("B")
                Write-Verbose "`tSearching for existing SSL Binding for $portNumber...";

                try {
                    Write-Verbose "`tAttempting to add SSL Binding..."
                    $output = Invoke-Command { "http delete sslcert ipport=0.0.0.0:$($portNumber)" | netsh } -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
                    Write-Verbose "`t`t$output"
                    $output = Invoke-Command { "http add sslcert ipport=0.0.0.0:$($portNumber) appid=$($guid) certhash=$($localhostCert.Thumbprint) certstorename=$($StoreName)" | netsh } -Verbose:$VerbosePreference -ErrorAction $ErrorActionPreference
                    Write-Verbose "`t`t$output"
                }
                catch [Exception]
                {
                    Write-Error $_
                }
            }
            Write-Verbose " "
        }
    }
}