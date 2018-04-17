<#
########################################
Script: InactiveUsersInitialRun.ps1
Author: Louis Crawley
Company: Trace Solutions
Date: 23/08/2017
Purpose: This script is to setup any references for the Inactive users script, and to make the inactive users script
         into a scheduled task
Notes: Please run as administrator
       This script can be run as many times as you want, it will replace any references with new information when run. 
########################################
#>

#Set execution policy to remotesigned
if((Get-ExecutionPolicy) -ne "RemoteSigned")
    {Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Force}
     

#If filepath to store txt file doesn't exist, make it.
$strFilePath = "C:\InactiveUsers\Exports\"
if(test-path $strFilePath)
    {}
else
    {New-Item -itemtype directory -force -path C:\InactiveUsers\Exports}



$DomainName = (Get-WmiObject win32_computersystem).domain

#Loads a vbnet inputbox for Domain Name
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$DomName = [Microsoft.VisualBasic.Interaction]::InputBox("Please type in your domain name (case-sensetive) - Suggested:" + "`r`n" + "       " + $domainname, "DOMAIN", "$DomName")

#Loads a vbnet inputbox for Admin Name
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$AdminName = [Microsoft.VisualBasic.Interaction]::InputBox("Please type in a user who's password won't change (Local System/ Local Service/ Network Service etc)", "Admin Username", "$AdminName")

#Loads a vbnet inputbox for Domain Name
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$AdminPassword = [Microsoft.VisualBasic.Interaction]::InputBox("Please type in the password for user:" + "`r`n" + "       " + $AdminName, "Admin Password", "$AdminPassword")

#Office365 Logon Credentials
$wshell = new-object -comobject wscript.shell
$wshell.popup("Please Enter Email Address/Pass for Account to Email From (Must have sufficient permissions to check licences",0,"Office365 Credentials",0x1)
$365Login = get-credential
$365LoginUser = $365Login.GetNetworkCredential().UserName
$365LoginPass = $365Login.GetNetworkCredential().Password

$FullUser = "$AdminName@$domainname"

#Input box for Email Addresses to fire the information to
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$EmailList1 = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the email address for the recipient of the script results (1 of 2):" + "`r`n" + "       " + $EmailList11, "Email Address 1", "$EmailList1")
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$EmailList2 = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter the email address for the recipient of the script results (2 of 2):" + "`r`n" + "       " + $EmailList22, "Email Address 2", "$EmailList2")

$EmailList = "$EmailList1 , $emaillist2"

[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$SMTPServ = [Microsoft.VisualBasic.Interaction]::InputBox("Please enter your SMTP Server Address. If you are having trouble, please go to mxtoolbox.com and enter in your domain to find." + "`r`n" + "       " + $SMTPServ1, "SMTP Address", "$SMTPServ")

#Create new text files referencing the Domain Name
New-Item C:\InactiveUsers\References\1\DOMAIN.txt -type file -force -value "$domname"
New-Item C:\InactiveUsers\References\2\1\Cred1.txt -type file -force -value "$FullUser"
New-Item C:\InactiveUsers\References\2\2\Cred2.txt -type file -force -value "$AdminPassword"
New-Item C:\InactiveUsers\References\3\1\Cred1.txt -type file -force -value "$365LoginUser"
New-Item C:\InactiveUsers\References\3\2\Cred2.txt -type file -force -value "$365LoginPass"
New-Item C:\InactiveUsers\References\4\MailingList.txt -type file -force -value "$EmailList"
New-Item C:\InactiveUsers\References\4\SMTP.txt -type file -force -value "$SMTPServ"

###Check if Scheduled Task Exists, if not continue

import-module scheduledtasks

$TaskName = 'Check_For_Inactive'

if( Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)
    {Unregister-ScheduledTask -TaskName $TaskName -confirm:$true}

#Create Scheduled Task#

#Get Script Location

$wshell.popup("Please select location of script for task",0,"Select Script Location",0x1)

$initialDirectory = "C:\"
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "PSScript (*.ps1,*.psm1,*.psd1,*.ps1xml,*.pssc,*.psrc)| *.ps1;*.psm1;*.psd1;*.ps1xml;*.pssc;*.psrc"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
#    write-host $openfiledialog -ForegroundColor Green
#    write-host $OpenFileDialog.SafeFileName -ForegroundColor Cyan
$FileDirec = $OpenFileDialog.FileName
$InactiveUsersScript = $OpenFileDialog.FileName

#$BatchContents = "powershell.exe -file " + '"' + "$inactiveusersscript" + '"'
#$BatchLocation = '"' + "C:\InactiveUsers\References\4\RunPS.bat"+ '"'
#New-Item C:\InactiveUsers\References\4\RunPS.bat -type file -force -value ("$batchcontents")

$ActionText = "powershell.exe"
$ArgumentText = "-executionpolicy bypass -file " + '"' + "$inactiveusersscript" + '"'

write-host $ActionText -ForegroundColor Yellow
$Action = New-ScheduledTaskAction "$actiontext" -Argument "$argumenttext"
$Trigger = New-ScheduledTaskTrigger -Weekly -At '1PM' -DaysOfWeek Monday
$Principal = new-scheduledtaskprincipal -UserId $365LoginUser -RunLevel Highest

$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings (New-ScheduledTaskSettingsSet) -Principal $principal


$Task | Register-ScheduledTask -TaskName $TaskName -User $fulluser -Password $AdminPassword

$wshell.popup("Scheduled Task setup. Please check 'Check_For_Inactive' in your schedular" + " `r`n" + "Script is finished, please close.")
