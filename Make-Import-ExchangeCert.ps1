############################################################################
#         Microsoft Exchange Certbot Certificate Replacement Script        #
#                 Written in December of 2022 by Jason Beaver              #
#                                                                          #
#  Prerequisites:  OpenSSL for Windows (x64), Microsoft Exchange, Certbot  #
#                     Operating System:  Windows Server                    #
############################################################################################################################
#  Insert into Task Scheduler to run once per week.                                                                        #
#  Suggested Parameters:                                                                                                   #
#  -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "C:\Windows\System32\Scripts\Make-Import-ExchangeCert.ps1" #
#                                                                                                                          #
#  Make a directory called "pfx" under C:\Certbot\live\<your domain here>\                                                 #
#  The user this runs as should have full permissions to that directory.                                                   #
#  MAKE SURE TO HAVE CERTBOT GENERATE AN RSA CERTIFICATE - EXCHANGE HATES OTHER TYPES!                                     #
############################################################################################################################
#   Use the log file for tracing if there are problems.
#$LogFile = "C:\users\username\Documents\certlog.txt"
#Start-Transcript -path $LogFile -append

$emailuser = "user"
$servername = "exchsvr"
$domainname = "example.com"

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
$installcert='yes'

#Compare current Exchange cert with the current ACME cert

$currentexchcert = Get-ExchangeCertificate | Select-Object Subject,Services,FriendlyName,NotAfter,Thumbprint | where { $_.FriendlyName -like "Exchange-*" -AND $_.Services -like "*IIS*"}

$File="C:\Certbot\live\" + $domainname + "\cert.pem"
$sslpath="C:\Program Files\OpenSSL-Win64\bin\openssl.exe"
$sslargs="x509 -in $File -dates -noout"

# Run OpenSSL and capture its output.
$p = New-Object System.Diagnostics.Process
$p.StartInfo.UseShellExecute = $false
$p.StartInfo.RedirectStandardOutput = $true
$p.StartInfo.FileName = $sslpath
$p.StartInfo.Arguments =  $sslargs
$p.Start() | Out-Null

# To avoid deadlocks, always read the output stream first and then wait.  
$SSLoutput = $p.StandardOutput.ReadToEnd()
$p.WaitForExit()

$SSLOutSplit = $SSLoutput.Split("`r`n")

$CertNotAfter1 = $SSLOutSplit[2]

#Strip off parts not needed
$CertNotAfter2 = $CertNotAfter1.TrimStart('notAfter=')
$CertNotAfter2 = $CertNotAfter2.TrimEnd(' GMT')

$format = "MMM dd HH:mm:ss yyyy"

# Since OpenSSL time is GMT, convert NotAfter to GMT
$ExchCertNotAfterDate = $currentexchcert.NotAfter.ToUniversalTime()

# Extract GMT time of ACME cert and parse as DateTime
$ACMECertNotAfterDate = [Datetime]::ParseExact($CertNotAfter2,$format,$null)

# Don't attempt to install a new cert if there's no new certificate
if ($ExchCertNotAfterDate -eq $ACMECertNotAfterDate){
    $installcert='no'
}

if ($installcert -eq 'yes')
{
$datetime = Get-Date -format 'MM-dd-yyyy HH:MM:ss'

$pfxdate = Get-Date -format 'MMddyyyy'

$certiisname="Exchange-"+ $pfxdate

# Make PFX file for import
$pfxname="C:\Certbot\live\" + $domainname + "\pfx\Exchange-"+ $pfxdate+".pfx"
$certname="C:\Certbot\live\" + $domainname + "\fullchain.pem"
$keyname="C:\Certbot\live\" + $domainname + "\privkey.pem"

$certpwd = "ChAng3Me"

# Make sure this points at your OpenSSL path
$sslpath="C:\Program Files\OpenSSL-Win64\bin\openssl.exe"

$sslargs="pkcs12 -export -password pass:" + $certpwd + " -out " + $pfxname + " -inkey " + $keyname + " -in " + $certname + " -name " + $certiisname

$p2 = New-Object System.Diagnostics.Process
$p2.StartInfo.UseShellExecute = $false
$p2.StartInfo.RedirectStandardOutput = $true
$p2.StartInfo.FileName = $sslpath
$p2.StartInfo.Arguments =  $sslargs
$p2.Start() | Out-Null

# To avoid deadlocks, always read the output stream first and then wait.  
$SSLoutput = $p2.StandardOutput.ReadToEnd()
$p2.WaitForExit()

# Check to make sure the PFX file was made.
$proceed = Test-Path -Path $pfxname

if ($proceed) {

$pfxpass = $certpwd | ConvertTo-SecureString -AsPlainText -Force

$newCert = Import-PfxCertificate -FilePath $pfxname `
    -CertStoreLocation "Cert:\LocalMachine\My" `
    -password $pfxpass

Import-Module Webadministration

    $sites = Get-ChildItem -Path IIS:\Sites

    foreach ($site in $sites)
    {
        foreach ($binding in $site.Bindings.Collection)
        {
            if ($binding.protocol -eq 'https')
            {
                $binding.AddSslCertificate($newCert.Thumbprint, "my")
            }
        }
    }
Enable-ExchangeCertificate -Thumbprint $newCert.Thumbprint -Services SMTP,POP,IMAP,IIS -DoNotRequireSsl -Confirm:$false -Force
Remove-ExchangeCertificate -Thumbprint $currentexchcert.Thumbprint -Confirm:$false

$smtpsvr = "server.example.com"
$smtpport = 25
$mailfrom = $servername + "@" + $domainname
$mailto = $emailuser + "@" + $domainname
$mailbody = "Certificate " + $certiisname + " installed on " + $servername + ".`nOpenSSL Output:`n" + $SSLoutput
$mailsubject = "Exchange Server Certificate " + $datetime
Send-MailMessage -port $smtpport -From $mailfrom -To $mailto -Subject $mailsubject -Body $mailbody -SmtpServer $smtpsvr
#Stop-Transcript
} else {
$smtpsvr = "server.example.com"
$smtpport = 25
$mailfrom = $servername + "@" + $domainname
$mailto = $emailuser + "@" + $domainname
$mailbody = "PFX file " + $pfxname + " wasn't there."
$mailsubject = "PFX File Not Found " + $datetime
Send-MailMessage -port $smtpport -From $mailfrom -To $mailto -Subject $mailsubject -Body $mailbody -SmtpServer $smtpsvr
#Stop-Transcript
}
} else {
$smtpsvr = "server.example.com"
$smtpport = 25
$mailfrom = $servername + "@" + $domainname
$mailto = $emailuser + "@" + $domainname
$mailbody = "No certificate installed on " + $servername + "."
$mailsubject = "No Exchange Server Certificate " + $datetime
Send-MailMessage -port $smtpport -From $mailfrom -To $mailto -Subject $mailsubject -Body $mailbody -SmtpServer $smtpsvr
}
#Stop-Transcript
