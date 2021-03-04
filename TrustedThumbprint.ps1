# This script adds a code signing certificate thumbprint to the list of trusted certificates
# for Microsoft Power BI connectors.
# Because it is a REG_MULTI_SZ registry value that Power BI reads, there aren't standard GPO mechanisms for
# handling such values so we resort to using a script.

param (
	[string]$SelfSigned = "false"
)

if ($SelfSigned -eq "false") {
	$Thumbprint = "1f1c059074762b49a97d6ce5c975a0a5e81bab5f"
} else {
	$certPassword = Read-Host "Code signing certificate password" -AsSecureString
	$pfxData = Get-PfxData -password $certPassword mobie\CodeSigning.pfx 
	$Thumbprint = $pfxData.EndEntityCertificates.Thumbprint
}

$RegKeyName = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Power BI Desktop'

if (!(Test-Path -Path $RegKeyName)) {
   New-Item -Path $RegKeyName
}

$RegKey = Get-Item $RegKeyName
$RegValue = $RegKey.GetValue("TrustedCertificateThumbprints")

if ($null -eq $RegValue) {
  New-ItemProperty -Path $RegKeyName -Name "TrustedCertificateThumbprints" -Value @( "$SelfSignedThumbprint" ) -PropertyType MultiString
} else {
  if (!($RegValue -contains "$SelfSignedThumbprint")) {
    $RegValue += "$SelfSignedThumbprint"
    Set-ItemProperty -Path $RegKeyName -Name "TrustedCertificateThumbprints" -Value $RegValue
  }
}