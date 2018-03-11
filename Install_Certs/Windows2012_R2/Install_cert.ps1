<#

.SYNOPSIS
This powershell script can be used to install a certificate (Root,Intermediate and Personal certs)

.DESCRIPTION
Tested platforms:
- Windows Server 2012R2 
- Windows 8.1

.Author
- Bharath Gajendran

.Version
1.0 - Added support for single and Multiple root and intermediate installtion
2.0 - Added support for Assigning the installed personal certificate to IIS website
3.0 - Added support for installing multiple personal certificates 
4.0 - Added support for installing p7b personal certificates
5.0 - Added support for installing PFX personal certificates
#>	

# Prerequisite check
if (-NOT([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
    Pause
    Throw "Administrator priviliges are required. Please restart this script with elevated rights."
}


#Installl Root certs placed in rootcerts directory
$Path0 = "$PSSCriptroot\rootcerts\"
$certFile = get-childitem $Path0  | where {$_.Extension -match "cer" -Or  $_.Extension -match "crt"}
$l1 = @(Get-ChildItem "$Path0" | where {$_.Extension -match "cer" -Or $_.Extension -match "crt"})
$l2 = $l1.count 
$i = 0
if (($certFile -ne $NULL) -and ($l2 -gt 1))
{
foreach ($cert in $certFile) 
   { 
     $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
     $cert.import($Path0 + $certfile.Name[$i])
     $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root","LocalMachine")
     $store.Open("MaxAllowed") 
     $store.Add($cert) 
     $store.Close()
     Write-Host "Certificate" $certfile.Name[$i] "- Installed SUCCESSFULLY!"
     $i++             
   }	
}
elseif (($certFile -ne $NULL) -and ($l2 -eq 1))
{
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName) 
    $store = get-item Cert:\LocalMachine\Root 
    $store.Open("ReadWrite") 
    $store.Add($cert) 
    $store.Close() 
	Write-Host "Certificate" $certfile.Name "- Installed SUCCESSFULLY!"
}
else
{
Write-Host "No Certs found in root certs directory"
}

#Install Intermediate certs pleaced in intermediate certs directory
$Path1 = "$PSSCriptroot\intermediatecerts\"
$certFile = get-childitem $Path1  | where {$_.Extension -match "cer" -Or  $_.Extension -match "crt"}
$l3 = @(Get-ChildItem "$Path1" | where {$_.Extension -match "cer" -Or  $_.Extension -match "crt"})
$l4 = $l3.count 
$j = 0
if (($certFile -ne $NULL) -and ($l4 -gt 1))
 {
 foreach ($cert in $certFile)
    {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.import($Path1 + $certfile.Name[$j])
        $store = get-item Cert:\LocalMachine\CA
        $store.Open("MaxAllowed") 
        $store.Add($cert) 
        $store.Close()
        Write-Host "Certificate" $certfile.Name[$j] "- Installed SUCCESSFULLY!"
        $j++ 
                
    }
 }	
elseif (($certFile -ne $NULL) -and ($l4 -eq 1))
{
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName) 
    $store = get-item Cert:\LocalMachine\CA 
    $store.Open("ReadWrite") 
    $store.Add($cert) 
    $store.Close() 
	Write-Host "Certificate" $certfile.Name "- Installed SUCCESSFULLY!"
}
else
{
Write-Host "No Certs found in intermediatecerts certs directory"
}

#Install personal certs with extension cer or crt pleaced in personal certs directory
$Path2 = "$PSSCriptroot\personalcerts\"
$location = "Cert:\LocalMachine\My"
$certFile = get-childitem $Path2  | where {$_.Extension -match "cer" -Or $_.Extension -match "crt"}
$l5 = @(Get-ChildItem "$Path2" | where {$_.Extension -match "cer" -Or $_.Extension -match "crt"})
$l6 = $l5.count
$m = 0
if (($certFile -ne $NULL) -and ($l6 -gt 1))
{
 foreach ($cert in $certFile)
    {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.import($Path2 + $certfile.Name[$m])
		$certhash = $cert.Thumbprint
		Import-certificate -Filepath ($Path2 + $certfile.Name[$m]) -CertStoreLocation $location >cert.txt
		$fname= Read-Host "Enter FriendlyName for the certificate" $certfile.Name[$m] "of your choice"
		(Get-ChildItem -Path $location\$certhash).FriendlyName = $fname
        Write-Host "Certificate" $certfile.Name[$m] "with friendly name" $fname "- Installed SUCCESSFULLY!" 
        Write-Host "Assigning SSL certificate" $certfile.Name[$m] "to the website"		
        Write-Host "Assigning SSL certificate to the default website"
        $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
        if ($pref -eq 'y')
        {
         $siteName = 'Default Web Site'
        }
        elseif ($pref -eq 'n')
        {
        $siteName = Read-Host "Enter the website name as shown in IIS manager"
        }
        else 
        {
        Write-Host "Invalid option enter y/n"
        }
        $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
        #set the ssl certificate
       $binding.AddSslCertificate($certhash, "my")
       Write-Host "Certificate Assigned successfully"
       $m++ 
    }
}	

elseif (($certFile -ne $NULL) -and ($l6 -eq 1))
 {
   $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName)
   $certhash = $cert.Thumbprint
   Import-certificate -Filepath $certFile.Fullname -CertStoreLocation $location >cert.txt
   $fname= Read-Host "Enter FriendlyName for the certificate" $certfile.Name "of your choice"
   (Get-ChildItem -Path $location\$certhash).FriendlyName = $fname
   Write-Host "Certificate" $certfile.Name "with friendly name" $fname "- Installed SUCCESSFULLY!"
   Write-Host "Assigning SSL certificate" $certfile.Name "to the website"
   $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
   if ($pref -eq 'y')
   {
    $siteName = 'Default Web Site'
   }
   elseif ($pref -eq 'n' )
   {
   $siteName = Read-Host "Enter the website name as shown in IIS manager"
   }
   else 
   {
   Write-Host "Invalid option enter y/n"
   }
   $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
   #set the ssl certificate
   $binding.AddSslCertificate($certhash, "my")
   Write-Host "Certificate Assigned successfully" 
}
else
{
Write-Host "No Certs of supported file type (CER or CRT) found in personal certs directory"
}

#Install personal certs with extension p7b pleaced in personal certs directory
$certFile = get-childitem $PATH2 | where {$_.Extension -match "p7b"}
$l7 = @(Get-ChildItem "$PATH2" | where {$_.Extension -match "p7b"})
$l8 = $l7.count 
$n = 0
if (($certFile -ne $NULL) -and ($l8 -gt 1))
 {
  foreach ($cert in $certFile)
  {
   Import-certificate -Filepath ($Path2 + $certfile.Name[$m]) -CertStoreLocation $location
   $certhash= Read-host "Enter the tumbprint of the certificate matching your URL (thumbprint for teh cert in P7B are shown above)"
   $fname= Read-Host "Enter FriendlyName for the certificate" $certfile.Name[$n] "of your choice"
   (Get-ChildItem -Path $location\$certhash).FriendlyName = $fname
   Write-Host "Certificate" $certfile.Name[$n] "with friendly name" $fname "- Installed SUCCESSFULLY!"
   Write-Host "Assigning SSL certificate" $certfile.Name[$n] "to the website"
   $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
   if ($pref -eq 'y')
   {
   $siteName = 'Default Web Site'
   }
   elseif ($pref -eq 'n')
   {
   $siteName = Read-Host "Enter the website name as shown in IIS manager"
   }
   else 
   {
   Write-Host "Invalid option enter y/n"
   }
   #getting site configuration
   $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
   #set the ssl certificate
   $binding.AddSslCertificate($certhash, "my")
   Write-Host "Certificate Assigned successfully"
   $n++
   }
  }
elseif (($certFile -ne $NULL) -and ($l8 -eq 1))
{
   Import-certificate -Filepath $certFile.Fullname -CertStoreLocation $location
   $certhash= Read-host "Enter the tumbprint of the certificate matching your URL (thumbprint for teh cert in P7B are shown above)"
   $fname= Read-Host "Enter FriendlyName for the certificate" $certfile.Name "of your choice"
   (Get-ChildItem -Path $location\$certhash).FriendlyName = $fname
   Write-Host "Certificate" $certFile.Name "with frinedly name" $fname "- IMPORTED SUCCESSFULLY!"
   $certhash = Read-host "Enter the thumbprint of the certificate as shown above:"
   Write-Host "Assigning installed SSL certificate" $certFile.Name "to the website"
   $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
   if ($pref -eq 'y')
   {
   $siteName = 'Default Web Site'
   }
   elseif ($pref -eq 'n')
   {
   $siteName = Read-Host "Enter the website name as shown in IIS manager"
   }
   else 
   {
   Write-Host "Invalid option enter y/n"
   }
   #getting site configuration
   $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
   $binding.AddSslCertificate($certhash, "my")
   Write-Host "Certificate Assigned successfully"
} 
else
{
Write-Host "No Certs of supported file type (P7B) found in personal certs directory"
} 

#Install personal certs with extension pfx pleaced in personal certs directory 
$certFile = get-childitem $PATH2 | where {$_.Extension -match "pfx"}
$l9 = @(Get-ChildItem "$PATH2" | where {$_.Extension -match "pfx"})
$l10 = $l9.count 
$p = 0
if (($certFile -ne $NULL) -and ($l10 -gt 1))
{
 foreach ($cert in $certFile)
    {
        $pfxpass= Read-Host "Enter the passoword of" $certfile.Name[$p] "to continue"
      	$pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $pfx.Import($Path2 + $certfile.Name[$p],$pfxpass,"Exportable,PersistKeySet") 
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store My,LocalMachine
        $store.Open('ReadWrite')
        $store.Add($pfx) 
        $store.Close() 
        $certhash = $pfx.Thumbprint
        Write-Host "Certificate" $certfile.Name[$p] "- IMPORTED SUCCESSFULLY!"    
        Write-Host "Assigning Installed SSL certificate" $certfile.Name[$p] "to the website"		
        Write-Host "Assigning SSL certificate to the default website"
        $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
        if ($pref -eq 'y')
        {
         $siteName = 'Default Web Site'
        }
        elseif ($pref -eq 'n')
        {
        $siteName = Read-Host "Enter the website name as shown in IIS manager"
        }
        else 
        {
        Write-Host "Invalid option enter y/n"
        }
        $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
        #set the ssl certificate
       $binding.AddSslCertificate($certhash, "my")
       Write-Host "Certificate Assigned successfully"
       $p++ 
    }
}	

elseif (($certFile -ne $NULL) -and ($l10 -eq 1))
 {
   $pfxpass= Read-Host "Enter the passoword of" $certFile.Name  "to continue"
   $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
   $pfx.Import($certFile.FullName,$pfxpass,"Exportable,PersistKeySet") 
   $store = New-Object System.Security.Cryptography.X509Certificates.X509Store My,LocalMachine 
   $store.Open('ReadWrite')
   $store.Add($pfx) 
   $store.Close() 
   $certhash = $pfx.Thumbprint
   Write-Host "Certificate" $certfile.Name "- IMPORTED SUCCESSFULLY!"
   Write-Host "Assigning SSL certificate" $certfile.Name "to the website"
   $pref = Read-Host "Do you want to assign certificate to Default website (y/n)"
   if ($pref -eq 'y')
   {
    $siteName = 'Default Web Site'
   }
   elseif ($pref -eq 'n')
   {
   $siteName = Read-Host "Enter the website name as shown in IIS manager"
   }
   else 
   {
   Write-Host "Invalid option enter y/n"
   }
   $binding = Get-WebBinding -Name "$siteName" -Protocol "https"
   #set the ssl certificate
   $binding.AddSslCertificate($certhash, "my")
   Write-Host "Certificate Assigned successfully" 
}
else
{
Write-Host "No Certs of supported file type (PFX) found in personal certs directory"
}

$tempfile= "$PSSCriptroot\cert.txt"
#delete temp file if exist
If (Test-Path $tempfile){
	Remove-Item $tempfile
}