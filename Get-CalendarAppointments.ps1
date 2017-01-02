###########################################################################
#Created By: Vijayan Kumaran
#Date: 28th Dec 2016
#Purpose: This script will Use EWS to get Appointments from Exchange Mailbox bsed on Start and End Date specified
#Usage: .\Get-CalendarAppointments.ps1 -Account account -Password password -Domain adatum.com -Identity Mailbox@adatum.com -StartDate 10-21-2017 -EndDate 10-22-2017
#Notes:
# - Accounts used in this script MUST have impersonation rights to the mailbox
# - Start and End Date format should be MM-DD-YYYY. You can specify the time as well for specific search. Eg: 10-21-2017 5:00PM 
# - You must install EWS API and update the DLL path in this script.
# - I only tested this script for Exchange 2010. You may test this script with Exchange 2013 and Exchange 2016. Before doing so, please change $Version variable value
###########################################################################


param(
	[parameter( Position=0, Mandatory=$true)]
		[string]$Account,
	[parameter( Position=1, Mandatory=$true)]
		[string]$Password,
	[parameter( Position=2, Mandatory=$true)]
		[string]$Domain,
        [parameter( Position=3, Mandatory=$true)]
		[string]$Identity,
        [parameter( Position=4, Mandatory=$true)]
		[string]$StartDate,
        [parameter( Position=5, Mandatory=$true)]
		[string]$EndDate
)

#Change this variable to your CAS server or OWA DNS Name
$EWSurl = "https://owa.outlook.com/EWS/Exchange.asmx"

Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010
$Forest = Set-ADServerSettings -ViewEntireForest $true

#Web Service
$EWSServicePath = "C:\Program Files\Microsoft\Exchange\Web Services\2.0\Microsoft.Exchange.WebServices.dll"
Import-Module $EWSServicePath

"$(Get-Date) starts...`r`n" > $LogPath

#Creating Service Object for Exchange

$Version = "Exchange2010_SP2"
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$Version
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $Account,$Password,$Domain


$Mailboxes = Get-Mailbox $Identity
foreach($Mailbox in $Mailboxes)
{
$MailboxToImpersonate = $Mailbox.WindowsEmailAddress.ToString()
Write-Host "Checking Calendar Appointment in - $MailboxToImpersonate`r`n" -ForegroundColor Green
"Checking Message Category for mailbox - $MailboxToImpersonate`r`n" >> $LogPath

$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate)
 
#Setting up EWS URL
$Uri=[system.URI] $EWSurl
$Service.URL = $Uri
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$MailboxToImpersonate)
 
#Setting up EWS URL
$Uri=[system.URI] $EWSurl
$Service.URL = $Uri

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxToImpersonate)
$CalendarFolder = [Microsoft.Exchange.WebServices.Data.CalendarFolder]::Bind($service,$folderid)
$cvCalendarview = new-object Microsoft.Exchange.WebServices.Data.CalendarView($StartDate,$EndDate,2000)
$cvCalendarview.PropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$frCalendarResult = $CalendarFolder.FindAppointments($cvCalendarview)
foreach ($apApointment in $frCalendarResult.Items){
     $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
     $apApointment.load($psPropset)
    "Appointment : " + $apApointment.Subject.ToString() 
    "Start : " + $apApointment.Start.ToString()
    "End : " + $apApointment.End.ToString()
    "Organizer : " + $apApointment.Organizer.ToString()
    "Required Attendees :"
    foreach($attendee in $apApointment.RequiredAttendees){
		"	" + $attendee.Address
	}
    "Optional Attendees :"
     foreach($attendee in $apApointment.OptionalAttendees){
		"	" + $attendee.Address
     }
    "Resources :"
     foreach($attendee in $apApointment.Resources){
		"	" + $attendee.Address
     }
    " "
}

}
