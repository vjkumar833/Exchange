###########################################################################
#Created By: Vijayan Kumaran
#Date: 8th Jan 2017
#Purpose: This script will Use Exchange Remote-Powershell to get Exchange Mailbox Folder permission for given User and send the report to Admin via Email
#Usage: .\Mailbox-FolderPermission.ps1 -Email Mailbox@adatum.com -Admin "Vijayan.Kumaran@Adatum.com"
#Notes:
# - I only tested this script for Exchange 2010. You may test this script with Exchange 2013 and Exchange 2016.
# - This script uses Exchange Remote Powershell. Kindly provide credentials that has Exchange permission for successful execution.
# - This script can be run from any domain-joined machine
###########################################################################

param(
[parameter( Position=0, Mandatory=$True)]
[string]$Email,

[parameter( Position=1, Mandatory=$True)]
[string]$Admin

)




#*************Variables*********************
$DIR = "D:\Inetpub\Script"
$Log = "$DIR\FolderPermission"
$Password = "Password"
$Domain = "Adatum.com"
$user = "Username"
$From = "NoReply@Adatum.com"
$SMTPServer = "mail.adatum.com"
$CASArray = "casarray.domain.com"
#*******************************************

################Exchange Remote Powershell################################
Set-ExecutionPolicy Unrestricted -Scope process
$pass = convertto-securestring "P$assword" -asplaintext -force
$creds = new-object -typename System.Management.Automation.PSCredential -argumentlist "$domain\$user",$pass
$connectionUri = 'https://$CASArray/powershell/?SerializationLevel=Full'
$session = New-PSSession -configurationname microsoft.exchange -connectionuri $connectionUri -Authentication Kerberos -AllowRedirection -Credential $creds
Import-PSSession `
-Session $session `
-WarningAction SilentlyContinue `
-ErrorAction SilentlyContinue `
-DisableNameChecking `
-OutVariable $Out
$Forest = Set-ADServerSettings -ViewEntireForest $true

##########################################################################

cd $DIR

out-File -FilePath "$Log\$Email.csv" -InputObject "Folder Name ; User ; Access Rights" -Encoding utf8

Write-Host "Checking Inbox permission for user $Email"
Write-Host ""
$Permissions = Get-MailboxFolderPermission $Email":\Inbox"
foreach ($item in $Permissions)
		{
		$users = $item.User
        $Rights = $item.AccessRights
        
        Foreach ($User in $users)
            {
             
            }
        Foreach ($Right in $Rights)
            {
             
            }
		Write-Host "$User has $Right Access to Inbox Folder" -Foregroundcolor "GREEN"
		out-File -FilePath "$Log\$Email.csv" -InputObject "Inbox; $User ; $Right" -Append -Encoding UTF8
		}
		Write-Host ""


$Folders = Get-MailboxFolderStatistics $Email | Where {$_.Identity -Like "$Email\Inbox\*"}
foreach ($items in $Folders)
{
$Folder = $items.Name
Write-Host "Checking $Folder permission for user $Email"
Write-Host ""

$Permissions = Get-MailboxFolderPermission $Email":\Inbox\"$folder | Where {$_.User -NotLike "Default*"}
foreach ($item in $Permissions)
		{
		$users = $item.User
        $Rights = $item.AccessRights
        
        Foreach ($User in $users)
            {
             
            }
        Foreach ($Right in $Rights)
            {
             
            }
		Write-Host "$User has $Right Access to $Folder Folder" -Foregroundcolor "Magenta"
		out-File -FilePath "$Log\$Email.csv" -InputObject "$Folder; $User ; $Right" -Append -Encoding UTF8
		}
}
		Write-Host ""

Start-Sleep 5


Write-Host "Sending Report to $Admin"

Send-MailMessage -Attachments "$Log\$Email.csv" -From $FROM -To $Admin -Body "Dear $Admin, <Br><Br> Attached is Mailbox Folder Permission Details for User: $Email. <br><br><br> Best Regards,<br>Email Support Team" -Subject "Mailbox Folder Permission Details for User: $Email" -SmtpServer $SMTPServer -BodyAsHtml