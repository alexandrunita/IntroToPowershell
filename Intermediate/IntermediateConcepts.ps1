# Powershell Providers
# A Windows PowerShell provider is basically a sort of abstraction layer that hides the complexities of different types of information stores.
Get-PSProvider
Get-ChildItem Env:
Get-Location
Set-Location -Path "HKLM:"
Set-Location -Path "Env:" -PassThru
Set-Location "C:\Users\vilega\OneDrive\Powershell"



# Discuss Like VS Match and Regex

<# !!! For the second presenation:
- need to review for, foreach, foreach-object
- need to review try,catch,finally
#>


#region Pass the value to the console but also re-use it
Get-mailbox admin2 |select * |Tee-Object -FilePath C:\Users\vilega\OneDrive\Powershell\Trainings\Me\Outputs\GetMailbox.txt 
Invoke-Item C:\Users\vilega\OneDrive\Powershell\Trainings\Me\Outputs\GetMailbox.txt

#endregion

#region Calculate Hash
# Consider if to include?
# Computes the hash value for a file by using a specified hash algorithm.
# Get-FileHash [-Path] <String[]> [-Algorithm <String> {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160} ] [ <CommonParameters>]

Get-FileHash C:\Users\vilega\Desktop\test20mb.file -Algorithm SHA1 | Format-List

#endregion


#region Send-MailMessage
# Consider if to include?
# !!! you need to have port 25 opened !!!
# [System.Net.ServicePointManager]::SecurityProtocol = 'Tls,TLS11,TLS12'

$time = Get-Date -Format yyyyMMdd_hhmmss
$creds = Get-Credential
Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -From admin@vilega05.onmicrosoft.com -To admin2@vilega05.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential $creds -Attachments "C:\Users\vilega\Desktop\test.com"
Send-MailMessage -SmtpServer vilega05.mail.protection.outlook.com -From admin@vilega05.onmicrosoft.com -To admin2@vilega05.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential -Attachments "C:\Users\vilega\Desktop\test.com"


#endregion


#region MessageTrace 
Get-MessageTrace |select -Last 1 | Get-MessageTraceDetail |fl

$msgID = (Get-MessageTrace |select -Last 1).MessageId
Get-MessageTrace -MessageId $msgID  | Get-MessageTraceDetail |fl

Get-MessageTrace -MessageId "56d6a6fc-e971-4d97-acc4-cf4f46e31d02@DB3FFO11FD054.protection.gbl"  | Get-MessageTraceDetail |fl
Get-MessageTrace | Out-GridView -PassThru |Get-MessageTraceDetail 
Get-MessageTrace | Out-GridView -PassThru |Get-MessageTraceDetail |fl

#endregion

#region Find MSOL Error
#Users with errors
Get-MsolUser -MaxResults 10000 -HasErrorsOnly

$UPN = "PFMBX1@vilega.onmicrosoft.com"
(Get-MSOLUser -UserPrincipalName  $UPN).errors.errorDetail.objectErrors.errorRecord.errorDescription 


#endregion


#region XML - Handson manipulation of XML Reports
$UPN = "PFMBX1@vilega.onmicrosoft.com"
$UPN = "admin@vilega.onmicrosoft.com"

# First Example - Depth:
$MSOLUserAll = Get-MsolUser -UserPrincipalName $UPN
$MSOLUserAll
$MSOLUserAll.Licenses
$MSOLUserAll.Licenses.ServiceStatus

Get-MsolUser -UserPrincipalName $UPN | Export-Clixml .\Outputs\MSOLUser.xml -Force
$MSOLUser = Import-Clixml .\Outputs\MSOLUser.xml
$MSOLUser
$MSOLUser.Licenses
$MSOLUser.Licenses.ServiceStatus

Get-MsolUser -UserPrincipalName $UPN | Export-Clixml .\Outputs\MSOLUserD.xml -Depth 4 -Force
$MSOLUserD = Import-Clixml .\Outputs\MSOLUserD.xml
$MSOLUserD
$MSOLUserD.Licenses
$MSOLUserD.Licenses.ServiceStatus


# Second Example -Migration Report
# Explore Migration XML, check why Depth > 2 does not give more info
# Collect from customer
$Mailbox = 'MigratedMailbox'
Get-MoveRequest
#Get-MoveRequestStatistics $Mailbox -IncludeReport -Diagnostic -DiagnosticArgument Verbose | Export-Clixml .\Outputs\MoveRequestStatistics.xml
Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" | Export-Clixml C:\Temp\MSSupport\MoveRequestStatistics_$Mailbox.xml 

# Analyze on the engineer side
$r = Import-Clixml .\Outputs\MoveRequestStatistics.xml
$r | fl | Out-File .\Outputs\MoveRequestStatistics.txt -Force; Invoke-Item .\Outputs\MoveRequestStatistics.txt

$r |fl 

$r | select * -ExcludeProperty DiagnosticInfo, Report

$r.Report
$r.Report.MailboxVerification | fl
$r.Report.MailboxVerification.missingitems | fl
$r.Report.Failures | select -Last 1

$r.Report.Failures.count
($r.Report.Failures | where {$_.FailureSide -like "*Source*"}).count

$r.Report.Failures | select -Last 1
$r.Report.Failures | ft -AutoSize Timestamp, FailureType, FailureSide
$r.Report.Failures | fl Timestamp, FailureType, FailureSide, Message

$r.Message
$r.OverallDuration
$r.PercentComplete
$r.Status
$r.StatusDetail
$r.Report.Entries
$r.Report.Entries.Failure


# Mailbox Import
Get-MailboxImportRequest -Mailbox admin |fl
Get-MailboxImportRequestStatistics -Identity (Get-MailboxImportRequest -Mailbox admin).RequestGuid -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" | Export-Clixml .\Outputs\ImportRequestStats.xml

$r2 = Import-Clixml .\Outputs\ImportRequestStats.xml
$r2 |fl
$r2 | select * -ExcludeProperty Report,DiagnosticInfo
$r2.Report.Failures 
$r2.Report.BadItems
$r2.Report.Connectivity
$r2.Report.Entries
$r2.Report.MailboxVerification
$r2.Report.SessionStatistics
$r2.Report.TargetMailboxSize


#endregion


# ==-=-==-=-=-=-=-=-=-=

trow
trow Exception("My exception")
trow RuntimeException

function TrapTest {
    trap {"Error found: $_"}
    thiswontwork
}


trap{ write-host $_; }
throw "blah"
write-host after

###

# from external, you'll get the code
$LASTEXITCODE

###

#region Function
Function Test
{
[CmdletBinding(SupportsShouldProcess)] # if error inside function, it will be visible outside the function
Param(
[Parameter(Mandatory=$True, HelpMessage = "Input your name", Position=0)]
[Alias("Numelemeu")]
[string] $name
)

$res = $PSCmdlet.ShouldContinue($name,"Title");
return $res
}
$newRes = Test -name "Victoras"
$newRes = Test -Numelemeu "Victoras"

#endregion 