<#

.SYNOPSIS
This powershell script can be used to generate a Certificate Signing Request (CSR) using the SHA256 signature algorithm and a 2048 bit key size (RSA).

Updated by - Bharath Gajendran

#>

####################
# Prerequisite check
####################

if (-NOT([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Administrator priviliges are required. Please restart this script with elevated rights." -ForegroundColor Red
    Pause
    Throw "Administrator priviliges are required. Please restart this script with elevated rights."
}


#######################
# Setting the variables
#######################
$UID = [guid]::NewGuid()
$files = @{}
$files['settings'] = "$($env:TEMP)\$($UID)-settings.inf";
$files['csr'] = "$($env:TEMP)\$($UID)-csr.req"


$request = @{}

$request['CN'] = Read-Host "Common Name (e.g. contoso.com)"
$request['O'] = Read-Host "Organisation (e.g. ABC Private Limited)"
$request['OU'] = Read-Host "Organisational Unit (e.g. Infra Services)"
$request['L'] = Read-Host "City (e.g. Plano)"
$request['S'] = Read-Host "State (e.g. Texas)"
$request['C'] = Read-Host "Country (e.g. US)"

#########################
# Create the settings.inf
#########################
$settingsInf = "
[Version] 
Signature=`"`$Windows NT`$ 
[NewRequest] 
KeyLength =  2048
Exportable = TRUE 
MachineKeySet = TRUE 
SMIME = FALSE
RequestType =  PKCS10 
ProviderName = `"Microsoft RSA SChannel Cryptographic Provider`" 
ProviderType =  12
HashAlgorithm = sha256
;Variables
Subject = `"CN={{CN}},OU={{OU}},O={{O}},L={{L}},S={{S}},C={{C}}`"
"
$settingsInf = $settingsInf.Replace("{{CN}}",$request['CN']).Replace("{{O}}",$request['O']).Replace("{{OU}}",$request['OU']).Replace("{{L}}",$request['L']).Replace("{{S}}",$request['S']).Replace("{{C}}",$request['C']).Replace("{{SAN}}",$request['SAN_string'])

# Save settings to file in temp
$settingsInf > $files['settings']

# Done, we can start with the CSR
Clear-Host

#################################
# CSR TIME
#################################

# Display summary
Write-Host "Certificate information
Common name: $($request['CN'])
Organisation: $($request['O'])
Organisational unit: $($request['OU'])
City: $($request['L'])
State: $($request['S'])
Country: $($request['C'])
Signature algorithm: SHA256
Key algorithm: RSA
Key size: 2048


" -ForegroundColor Yellow

certreq -new $files['settings'] $files['csr'] > $null 

# Output the CSR
$CSR = Get-Content $files['csr']
Write-Output $CSR
Write-Host "
"

# Set the Clipboard (Optional)
Write-Host "Copy CSR to clipboard? (y|n): " -ForegroundColor Yellow -NoNewline
if ((Read-Host) -ieq "y") {
	$csr | clip
	Write-Host "Copied to the clipboard
"
}

# Save CSR to CSRFolder
$CSR = Get-Content $files['csr'] | out-file -filepath $PSSCriptroot\CSRFolder\$($request['CN']).txt
Write-Host "CSR is also saved under CSRFolder"

########################
# Remove temporary files
########################
$files.Values | ForEach-Object {
    Remove-Item $_ -ErrorAction SilentlyContinue
}
