#use this to install, uncomment and install
Install-Module ExchangeOnlineManagement

#use this to connect and uncomment the connect
Connect-ExchangeOnline

Import-Module ExchangeOnlineManagement

$mailboxes = Get-Mailbox | Select-Object -Property Name

$thina = @()

foreach($mail in $mailboxes)
{


$thing = $mail -replace "@{Name=" -replace ""

$thing2 = $thing -replace "}" -replace ""


$val = Get-InboxRule -Mailbox $thing2 -ErrorAction SilentlyContinue | Select-Object -Property Name, Enabled

$yourData = @(
    [pscustomobject]@{Email=$thing2; Name='' ; Enabled=''}
    [pscustomobject]$val
    
    ) 


    Write-Host $yourData
 
 $thina += $yourData 





}


$thina | Export-Csv -Path C:\MailRule.csv

