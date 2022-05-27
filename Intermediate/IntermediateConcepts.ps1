#region Loops

# Conditional Logic - (if, elseif, else, switch)
if (condition) {code block}
elseif (condition) {code block}
else  {code block}

[int]$num = Read-Host "Please input an integer number"
if ($num -gt 0) {Write-Host "Pozitive number"}
elseif ($num -lt 0) {Write-Host "Negative number"}
else  {Write-Host "This number is zero"}


Switch (<testvalue>)
{
    <condition1> {<action>}
    <condition2> {<action>; break}
    <condition3> {<action>}
    <condition4> {<action>}
    <condition5> {<action>}
    default {<action>}
}

$switchTest=4
Switch ($switchTest)
{
    1 {Write-Host "1"; }
    2 {Write-Host "2"; }
    3 {Write-Host "3"; }
    default {Write-Host "default"}
}
# !!! The "Default" keyword specifies a condition that is evaluated only when no other conditions match the value.

switch (3)
{
    1 {"It is one."}
    2 {"It is two."}
    3 {"It is three."; Break}
    4 {"It is four."}
    3 {"Three again."}
}
It is three.

# !!! Any Break statements apply to the collection, not to each value.
switch (4, 2)
{
    1 {"It is one."; Break}
    2 {"It is two." ; Break }
    3 {"It is three." ; Break }
    4 {"It is four." ; Break }
    3 {"Three again."}
}

It is four.

<# Conditional Logic - Loops

while - The condition is evaluated first, and if true it starts running the code block
while
while (condition) {code block}


do while - Script block executes as long as condition value = True.
do while
do {code block}
while (condition)

do until – Script block executes until the condition value = True.
do {code block}
until (condition)

'do while' and 'do until' runs first the code block then evaluates the condition
#>

# Example 'While':

$i = 0
While ($true) {
    Write-Host "True"
    $i++
    if ($i -eq 3) {break}
}

# Example 'Do Until':

Write-Host "Type 'Yes' proceed or 'No' to stop and exit"   
    Do {
        [string]$option = read-host "Please type 'Yes' proceed or 'No' to stop and exit"
        $option = $option.ToLower()
    } Until (($option -eq "yes") -or ($option -eq "no"))
        
    If ($Option -eq "yes") { 
        Write-Host -ForegroundColor Yellow "Proceeding..."
    }
    Else {
        Write-Host -ForegroundColor Yellow "Exitting..."
}


# Example 'Do While':

Write-Host "Type 'Yes' proceed or 'No' to stop and exit"   
    Do {
        [string]$option = read-host "Please type 'Yes' proceed or 'No' to stop and exit"
        $option = $option.ToLower()
    } While (!(($option -eq "yes") -or ($option -eq "no")))
        
    If ($Option -eq "yes") { 
        Write-Host -ForegroundColor Yellow "Proceeding..."
    }
    Else {
        Write-Host -ForegroundColor Yellow "Exitting..."
}   
        

<#

for – Script block executes a specified number of times.
for (initialization; condition; repeat/increment/decrement)
{code block}

foreach - Executes script block for each item in a collection or array.
foreach ($<item> in $<collection>)
{code block}


ForEach-Object - Executes script block for each item in a collection or array but is used after pipeline ($_)
$<collection> |  ForEach-Object { code block}
#>

for($i=0; $i -lt 10;$i++){write-host $i} # loop 10 times

$i=0; for(;;){Write-Host $i; $i++; if($i -eq 10){break;}} #simulates a while(true) loop with a condition to break out of the loop

# foreach - keyword
$mbxs = Get-Mailbox
foreach ($mbx in $mbxs){$mbx.alias}

### !!!!!!!!! !!!!!!!!!!!!!!!!!!!!

# foreach alias for Foreach-Object cmdlet
$mbxs | foreach {$_}
get-mailbox | foreach {$_}

For($i=0;$i -le 4;$i++) {
    write-host($i);
}

$mbxs = get-mailbox
$mbxs.GetType()
$mbxs[0]
foreach ($mbx in $mbxs)
{
    $mbx.alias
}

$mbxs | ForEach-Object {$_.alias}

for ($i=0;$i -lt ($mbxs.length); $i++)
{
    $mbxs[$i].alias
}

#endregion

#region Errors
# Muliple ways to do it
# To keep it simple you can use bellow command to see last error details
$error[0].Exception | fl * -Force

$e = $error[0];
$e.Exception | fl * -f
$e.Exception.SerializedRemoteException | fl * -f

# from PS you'll get true/false depending on the result of last cmdlet
$?

# try / catch
$ErrorActionPreference = "Continue"
$ErrorActionPreference = 'Stop'

try { get-mailbox admin -Erroraction Stop}
catch{ Write-Host "Something went wrong..."}
finally {write-host "Regardless..."}

#endregion

#region to add a new value, keeping old values on a property (to add/remove email address, domains, IPs, etc...)
Get-Mailbox user8 |fl *email*
# To add v1
$mbx =  (Get-Mailbox user8).EmailAddresses
$mbx.Add("smtp:user8abc@axul.onmicrosoft.com")
Set-Mailbox user8 -EmailAddresses $mbx
Get-Mailbox user8 |fl *email*

# To remove v1
$mbx =  (Get-Mailbox user8).EmailAddresses
$mbx
$mbx.Remove("smtp:user8abc@axul.onmicrosoft.com")
$mbx
Set-Mailbox user8 -EmailAddresses $mbx
Get-Mailbox user8 |fl *email*

# To add / remove v2
Get-Mailbox user8 |fl *email*
Set-Mailbox user8 -EmailAddresses @{add="smtp:user8abc@axul.onmicrosoft.com"}
Get-Mailbox user8 |fl *email*
Set-Mailbox user8 -EmailAddresses @{remove="smtp:user8abc@axul.onmicrosoft.com"}
Get-Mailbox user8 |fl *email*

#endregion



#region Array/HashTable
# @() - Array - data collection, indexable, immutable
# @{} - HashTable - key value collection (index - unsorted key value pair)
# $hashtable = @{name="Victor";age="36"}

$collection= @()
$object1 = New-Object PSObject
$object2 = New-Object PSObject
$object1 | Add-Member -Name Name -Value "Victor" -MemberType NoteProperty
$object2 | Add-Member -Name Name -Value "Andrei" -MemberType NoteProperty 
$collection+= $object1
$collection+= $object2
$collection | foreach{Write-host($($_.Name))}

#endregion

#region Split
# dotnet embedded function to split string

$a = "this is a sample text"
$a.Split(" ")[0]

$a | Get-Member
$a.ToUpper()

#endregion 

#region Remember
<#
- Everything is an object
- Check if Customer is having at least minimum required version of PowerShell (example: 3)
- Start-Transcript
- history
- Run "$FormatEnumerationLimit = -1" before collecting formatted information
- If you check all objects please user "-Resultsize Unlimited", or "-All" depends on the cmdlet used
- Always filter on the left side of the pipeline and format only on the right side pipeline
- Always run the cmdlet on your side before sending them to customers!!
- We do not provide scripts, but if we do, it is as a best effort and they are to be considered as sample scripts for the customer insipire from, when creating their own. Such sample scripts must be accompaied by our disclaimer:

    This is a sample script and sample scripts are not supported under any Microsoft standard support program or service. The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

- Run the cmdlets first with -WhatIf
- The most important cmdlets are:
    Get-Help
    Get-Command
    Get-Member
- When you have a complex situation, break it in small pieces that can be manageable  and tested
- !!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.
#>


#endregion



"10-20" -Contains "-"
"10-20".Contains("-")
"10-20" |gm

"10-20" -split "" -contains "-"

@("10","20","30") -contains "20"


get-help System.Collections.ArrayList

# get more ingo on hashtables
Get-Help about_Hash_Tables 
get-help ConvertFrom-StringData
ConvertFrom-StringData
$p.keys | foreach {$p.$_.handles}

# Optimizing PowerShell Scripts
Start-Process "https://blogs.technet.microsoft.com/ashleymcglone/2017/07/12/slow-code-top-5-ways-to-make-your-powershell-scripts-run-faster/"
Start-Process "https://blogs.technet.microsoft.com/heyscriptingguy/2014/05/18/weekend-scripter-powershell-speed-improvement-techniques/"
Start-Process "https://social.technet.microsoft.com/wiki/contents/articles/11311.powershell-optimization-and-performance-testing.aspx"
Start-Process "https://blogs.technet.microsoft.com/heyscriptingguy/2014/05/17/weekend-scripter-best-practices-for-powershell-scripting-in-shared-environment/"
Start-Process "https://blogs.msdn.microsoft.com/powershell/2008/01/28/lightweight-performance-testing-with-powershell/"


# Foreach statment vs foreach alias (foreach-object)
Start-Process "https://poshoholic.com/2007/08/21/essential-powershell-understanding-foreach/"
Start-Process "https://poshoholic.com/2007/08/31/essential-powershell-understanding-foreach-addendum/"

$block1={
get-mailbox |foreach {$_.alias}
}

$block2={
foreach ($mb in get-CASmailbox ){$_.alias}
}

(Measure-Command $block1).TotalMilliseconds
(Measure-Command $block2).TotalMilliseconds

foreach ($character in [char[]]"aeioubcd") { if (@('a','e','i','o','u') -contains $character ) { continue } $character }
[char[]]"aeioubcd" | foreach { if (@('a','e','i','o','u') -contains $_ ) { continue } $_ }


#Begin
$path=[Environment]::GetFolderPath("Desktop")
#$path = "c:\temp" 
$timestamp = Get-Date -format yyMMdd_hhmmss
Start-Transcript -Path "$Path\Transcript_$timestamp.txt" -Force

Stop-transcript
#End


$string1 = $null
IF ([string]::IsNullOrWhitespace($string1)){'empty'} else {'not empty'}


Get-OrganizationConfig |fl *block*
help Set-OrganizationConfig -Parameter IPListBlocked


$calendars = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Get-MailboxFolderStatistics | ? {$_.FolderType -eq "Calendar"} | select @{n="Identity"; e={$_.Identity.Replace("\",":\")}}
$calendars | % {if ((Get-MailboxFolderPermission -Identity $_.Identity -User Default).AccessRights -ne "Reviewer") {Set-MailboxFolderPermission -Identity $_.Identity -User Default -AccessRights Reviewer}} 

$calendars | % {Set-MailboxFolderPermission -Identity $_.Identity -User Default -AccessRights LimitedDetails} 



[string]$myvar2 = 1
$myvar2.GetType()
$myvar1 = 1
$myvar1.GetType()


@( “machine1” , “machine2” , “machine3”).GetType()
$arr = @( “machine1” , “machine2” , “machine3”)
$arr | Get-Member
$arr[0] | Get-Member
$arr[0].GetType()
$arr.GetType()




New-MoveRequest user8
Get-MoveRequestStatistics user88 | fl status
Do {
        $stats = Get-MoveRequestStatistics -Identity "user88"
        Write-Host '*' -NoNewline
        Start-Sleep -Seconds 5
    } While ($stats.Status -ne 'Completed')
    Write-Host "`n> Move request completed"



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

# region - TO REIEW


$AllMbx = get-mailbox 

$AllMbx.GetType()
$mbx= get-mailbox test4
$AllMbx.Contains('test4')
([System.Collections.ArrayList]$AllMbx).Contains($mbx)
$AllMbx.Contains($mbx)

$mbx=$AllMbx[0]
$AllMbx.Contains($mbx)
Start-Process "http://www.computerperformance.co.uk/powershell/powershell_conditional_operators.htm"


# endregion