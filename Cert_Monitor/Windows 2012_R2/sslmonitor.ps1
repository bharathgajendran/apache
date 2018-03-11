<#

.SYNOPSIS
This powershell script can be used to monitor a certificate that are used by the websites running on IIS.

.DESCRIPTION
Tested platforms:
- Windows Server 2012R2 
- Windows 8.1

.Author
- Bharath Gajendran

.Version
1.0 - Intial Version
2.0 - Modified script to send email only if certiifcate match is found.
3.0 - Modified the email content for better notification.
4.0 - Logging fetaure enabled.
#>	


Import-module webadministration

function Get-TimeStamp {
    
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    
}
$DaysToExpiration = 700
$expirationDate = (Get-Date).AddDays($DaysToExpiration)
$report = @()
$sites = Get-Website | ? { $_.State -eq "Started" } | % { $_.Name }
foreach ($site in $sites){
    #Write-host $site.ToUpper() -ForegroundColor Yellow
    $certs = Get-ChildItem IIS:\SslBindings | ? {
           $site -contains $_.Sites.Value
         } | % { $_.Thumbprint }
         Foreach ($cert in $certs){
            $line = Get-ChildItem Cert:\LocalMachine\My | Where {$_.Thumbprint -eq $Cert -and $_.NotAfter -lt $expirationDate}  | Select @{N="SiteName";E={$site}},DnsNameList, @{N="Subject Alternative Name";E={($_.extensions | where {$_.Oid.friendlyname -eq "Subject Alternative Name"}).format(1)}}, NotAfter, Issuer
            $report += $line
         }
}
$report | Export-Csv $PSSCriptroot\$env:COMPUTERNAME.csv -NoTypeInformation -UseCulture

$check = Import-Csv "$PSSCriptroot\$env:COMPUTERNAME.csv" 

if ($check) 
{  
$htmlformat  = '<title>Table</title>'
$htmlformat += '<style type="text/css">'
$htmlformat += 'BODY{background-color:#FFFFFF;color:#000000;font-family:Arial Narrow,sans-serif;font-size:17px;}'
$htmlformat += 'TABLE{border-width: 3px;border-style: solid;border-color: black;border-collapse: collapse;}'
$htmlformat += 'TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:#FFFFFF}'
$htmlformat += 'TD{border-width: 1px;padding: 8px;border-style: solid;border-color: black;background-color:##FFFFFF}'
$htmlformat += '</style>'
$bodyformat = '<h1>Table</h1>'

$certsInfo = Import-Csv -Path $PSSCriptroot\$env:COMPUTERNAME.csv | ConvertTo-Html -Head $htmlformat 

$mailBody = 
@"
Hello,</br>
</br>
Found the following certificates used in $env:COMPUTERNAME server are expiring shortly.</br>
</br>
$certsInfo
</br>
Kindly do the needful to get it renewed.</br>
</br>
Regards,</br>
SSL Monitoring.</br>
"@

Send-MailMessage -Body $mailBody -BodyAsHtml `
-From "$env:COMPUTERNAME@toyota.com" -To "emailaddress.com" `
-Subject "SSL Certificate expiry notification" -Encoding $([System.Text.Encoding]::UTF8) `
-SmtpServer "xxxxxxxx"

Write-Output "$(Get-TimeStamp) Task completed and email sent successfully." | Out-file $PSSCriptroot\$env:COMPUTERNAME.sslmonitor.log -append

}
else 
{
Write-Output "$(Get-TimeStamp) Task completed successfully and no certificates nearing expiry are found." | Out-file $PSSCriptroot\$env:COMPUTERNAME.sslmonitor.log -append
}
