<#
#####################################
Script: CheckForInactiveUsers.ps1
Author: Louis Crawley
Company: Trace Solutions
Date: 23/08/2017
Purpose: This script is to check for inactive users in the AD, list their name, last active date, O365 Licence type
         in a csv file and send it off via email.
Notes:
#####################################
#>

#If filepath to store file doesn't exist, make it.
$strFilePath = "C:\InactiveUsers"
if(test-path $strFilePath)
    {}
else
    {
        $wshell = new-object -comobject wscript.shell
        $wshell.popup("You need to run the file: InactiveUsersInitialRun.ps1 before this can be used",0,"Requires References",0x1)
        exit
    }

import-module activedirectory

#Get date for the file naming
$dt = get-date -format ddMMyyyy

#Get Domain from Refs
$DOMAIN = get-content C:\InactiveUsers\References\1\DOMAIN.txt

#Office 365 Connectivity
#Create O365Credentials From Refs and store in $365Credentials variable as PSCredential
$365User = get-content C:\InactiveUsers\References\3\1\Cred1.txt
$365Pass = get-content C:\InactiveUsers\References\3\2\Cred2.txt
#$365Pass = "Puga9131"
$365Pass = ConvertTo-SecureString "$365Pass" -AsPlainText -force
$365Credentials = New-Object system.management.automation.pscredential($365User,$365Pass)

write-host $365pass -ForegroundColor Yellow
###
#Use O365 Credentials to connect
import-module msonline
$365PSSession = New-PSSession -ConfigurationName microsoft.exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $365Credentials -Authentication basic -AllowRedirection
Import-PSSession $365PSSession -AllowClobber
Connect-MsolService -Credential $365Credentials

#Set how long to check for users being inactive
$InactiveDays = new-timespan -days 30
#Setup Exempt Group from AD Search
#$ExemptionGroup = get-adgroup ".TPS Blue Team" | select CN | out-string
$SEARCHAD = search-adaccount -usersonly -accountinactive -timespan $InactiveDays | get-aduser | Where-Object {$_.Enabled -eq $true} | select-object name
$InactiveUsers = $SEARCHAD | Out-String
$Inactiveusers = $InactiveUsers.Split("`n",[system.stringsplitoptions]::RemoveEmptyEntries)

#Couple of static variables
$TS = "Solutions"
$SMTPServ = get-content C:\InactiveUsers\References\4\SMTP.txt
#$TPS = 'Payroll'
#$SolutionsMember = get-adgroupmember -Identity $TS -Recursive | select -ExpandProperty name
#$PayrollMember = get-adgroupmember -Identity "*$TPS*" -Recursive | select -ExpandProperty name

#Loop through the Inactive AD users, and query the O365 account for the licence type associated with the user then assign to $ExportingItems variable
#write-host $InactiveUsers -ForegroundColor Yellow

$ExportingItems = ''
$ExportingItems = @()

ForEach($line in $InactiveUsers)
{
    $Licence = ""
    $LicenceType = ""
    $InactiveTime = ""

    if ($line -ne $null)
    {
        $Licence = (Get-MsolUser -SearchString $line).licenses.servicestatus | Out-String -ErrorAction SilentlyContinue
            if($Licence.Contains("EXCHANGE_S_ENTERPRISE"))
            {
                $LicenceType = "E3 O365 Licence"
            }
            elseif($Licence.Contains("EXCHANGE_S_STANDARD"))
            {
                $LicenceType = "E1 O365 Licence"
            }
            else
            {
                $LicenceType = "Unlicenced"
            }
    }
    else
    {
        $LicenceType = "Unlicenced"
    }

    
<#   if ($SolutionsMember -contains $line)
        {
        $GroupName = 'Solutions'
        }
    elseif ($PayrollMember -contains $line)
        {
        $GroupName = 'Payroll Services'
        }
    else
        {
        $GroupName = 'Other'
        }
#>


    $NameForExport = $line
    $LicenceForExport = $LicenceType

    $ExportingItems += new-object psobject -Property @{
                                    Name = $NameForExport
                                    LicenceType = $LicenceForExport
                                    #Company = $GroupName
                                                      }


}


#write-host 'Please let this work ლ(ಥ Д ಥ )ლ otherwise I will kill myself (つ◉益◉)つ ' -ForegroundColor Yellow

#Change the $ExportingItems variable to an object, and export to csv file
$ExportingItems = $ExportingItems | Select-Object Name, LicenceType | Sort-Object -property licencetype, name #, GroupName
$ExportingItems | export-csv "C:\InactiveUsers\Exports\Inactive_Users_$dt.csv" -notypeinformation


#Write-Host 'Why you no work script!?!? └༼ ಥ ᗜ ಥ ༽┘ └༼ ಥ ᗜ ಥ ༽┘ └༼ ಥ ᗜ ಥ ༽┘' -ForegroundColor Red

#Email CSV file to relevent people.


################ TEST #####################

#$Mailbox = get-mailbox -Identity 'Louis Crawley' | out-string
#write-host $Mailbox -ForegroundColor Cyan

#$Test = $Mailbox.name
#write-host $Test -ForegroundColor Yellow

$MailingList = get-content C:\InactiveUsers\References\4\MailingList.txt


#Set a variable with the time of day for email body text
if( (get-date -uformat %p) -eq "AM") {
    $Time = "Morning"
    }
else {
    $Time = "Afternoon"
    }

$Attachment = "C:\InactiveUsers\Exports\Inactive_Users_$dt.csv"


#Does not work when running from Powershell command line, only in ISE

#Create an outlook session for emailing
add-type -assembly "Microsoft.Office.Interop.Outlook"
$ol = new-object -ComObject outlook.application
$ns = $ol.getnamespace("MAPI")

$BodyText = "Good $time," + "`r`n" + "`r`n" + "Please find attached a current list of Inactive AD users on the " + "'" + "$DOMAIN" + "'" + " domain, and their respective Office365 licences." + "`r`n" + "`r`n" + "Regards," + "`r`n" + "`r`n" + "TST"


$mail = $ol.createitem(0)
$mail.to = $mail.recipients.add($MailingList)
$mail.subject = "Current Inactive Users - $DOMAIN"
$mail.body = $BodyText
$mail.Attachments.add($Attachment)
$mail.Send()

#TEST

#$inspector = $mail.getinspector
#$inspector.activate()



<#
write-host $365user


$Msg = "Good $time," + "`r`n" + "`r`n" + "Please find attached a current list of Inactive AD users on the " + "'" + "$DOMAIN" + "'" + " domain, and their respective Office365 licences." + "`r`n" + "`r`n" + "Kind Regards," + "`r`n" + "`r`n" + "Technical Services Team" + "`r`n" + "Trace Solutions - Part of Trace Group"

Send-MailMessage -To $MailingList -from $365User -SmtpServer "$SMTPServ" -Subject "Current Inactive Users - $domain" -Attachments $attachment -body $msg -Credential $365Credentials
#>


#Eliminate unlicenced users, and blank user names from the end result.