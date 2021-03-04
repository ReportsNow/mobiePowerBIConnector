$certName = read-host "Enter name for code signing certificate (default: ""$env:ComputerName PowerBI"")"
if ($certName -eq "")
{
	$certName = "$env:ComputerName PowerBI"
}

$cert = New-SelfSignedCertificate -DnsName $certName -Type CodeSigning -CertStoreLocation Cert:\CurrentUser\My

$CertPassword = read-host "Certificate Password" -AsSecureString
Export-PfxCertificate -Cert "cert:\CurrentUser\My\$($cert.Thumbprint)" -FilePath "mobie\CodeSigning.pfx" -Password $CertPassword