#region WhatIsPowerShell
Write-Host "Windows PowerShell is an interactive object-oriented command environment with scripting language features that utilizes small programs called cmdlets to simplify configuration, administration, and management of heterogeneous environments in both standalone and networked typologies by utilizing standards-based remoting protocols."
<# Powershell editors:
- PowerShell
- Integrated Scripting Environment (ISE)
- 3rd party editors 
#>
#endregion

<#
    We will not go into Workload specifics here.
    Each M365 Workload has its own Powershell Connection steps and cmdlets documented in official articles.
    We will first focus on fundamentals of Powershell use on Windows.
#>

#region Syntax

<# 
Verb-Noun -ParameterName ParameterValue

Verbs (Get-Verb)
    - New (create)
    - Get
    - Set
    - Remove
    - Show
    - Write
    - Clear
    - Get
    etc.

Full List of approved verbs and definition : https://docs.microsoft.com/en-us/powershell/scripting/developer/cmdlet/approved-verbs-for-windows-powershell-commands?view=powershell-7.2

# Comments
# comment a row

<#  

This is a comment block
   
block comment

#>

#endregion Syntax

#region Aliases
Get-Alias history
Get-Alias dir
Get-Alias |select -First 10

# For waht is cls alias?
Get-Alias cls

# For waht is dir alias?
Get-Alias dir

# For waht is cd alias?
Get-Alias cd

#endregion

#region Check Environment Settings

# Check Powershell version
$PSVersionTable

# Console Settings
Get-Host

# Local Machine settings
Get-Culture
# note the culture that PowerShell is using (in this case en-US)

#If you need to run your PowerShell in a lower version, open Powershell with the following command: "powershell.exe -version 2"
#In this case Powershell version used was 2
#endregion

#region HelpTool

# first thing to do on a new machine
# Update Help 
Update-Help

# If we want to save Help content to a file, run:
Save-Help

# We can also update-help from a saved file via save-help if without Internet Connection
Update-Help -SourcePath %%

# Use powershell help first instead of any search engine
#Default command
get-help Get-Mailbox #If not connected to Workload Module, help will not retrieve the information
# To only retrieve help on a specific parameter:
get-help Get-Service -Parameter name

# to see the examples:
get-help Set-Mailbox -examples

# for more information
get-help Set-Mailbox -detailed

# for technical information
get-help Get-Mailbox -Parameter Identity

# for online help 
get-help Set-Mailbox -online

# to explore all options of Get-Help cmdlet, check the help repository:
Get-Help Get-Help

# Help structured on topics (like user manual for powershell). Recommend to use it !!!
# We can use alias "help" for Get-Help cmdlet
help about

# examples
help about*
help about_Operators
help about_Aliases
help about_Arithmetic_Operators
#endregion HelpTool

#region Get-Command
# List of the commands available. You can use wildcards to find it. (the results will be:  cmdlets, functions, workflows, and aliases )
# Alias for "Get-Command" is "gcm"
# Details retrived from local module/dll
# Will not retrieve information if module is not loaded, except for installed module ExportedCommands

Get-Command Get*MSOL*
Get-Command *Calendar*

Get-Command -ParameterName max*
Get-Command -ParameterName UserPrincipalName

# Get-Member (gm alias)
#Any command that produces output on the screen can be piped to Get-Member in order to see the events, alias properties, methods, properties and note properties
Get-Mailbox admin | Get-Member
(Get-Mailbox admin).gettype()

# If you need more info : help Get-Command -online/-detailed/-full/etc...
#endregion Get-Command

#region Profile

# Profiles (help about_Profiles)
# You can create a Windows PowerShell profile to customize your environment and to add session-specific elements to every Windows PowerShell session  that you start.
# For example, the $Profile variable has the following values in the Windows  PowerShell console.
$Profile 
$Profile.AllUsersCurrentHost
$Profile.CurrentUserAllHosts 
$Profile.AllUsersAllHosts

# Each Powershell Console will have its own profile config file in same folder paths, for example ISE
# For example, you could of override Powershell default prompt function using your profile
# function prompt {"$( ((Get-Location).Path).Split("\")|select -last 1 )>"}

#endregion Profile

#region Parameters
# Positional parameters make Windows PowerShell commands shorter because you do not have to do as much typing,
# and some of the parameter names are rather long. How to find positional parameters?
help Get-Mailbox -Full

# or if you know it but whant to check
help Get-Mailbox -Parameter Identity

# Partial Parameters (shortcuts)
Get-Mailbox -Id admin

<#
If the cmdlet is making changes you can use the following parameters:
- WhatIf (to see what will be changed)
-Confirm:$False (in case a permission need to be asked, it will automatically proceed)
#>

# We recommend avoiding to use positional parameters when writing scripts/collaborating with others as it can make code harder to read
Write-Host "!!! In a script, you should always use complete parameter names because with full parameter names, you actually know what the code is doing and it is more readable." -ForegroundColor "Red"
Write-Host "!!! You should avoid partial parameter names, and positional parameters for the same reason." -ForegroundColor "Red"
#endregion

#region History
# History of last commands
History
history | select -First 1 | fl *
History | select CommandLine
history | select -last 10
#endregion

#region Snapin and Modules
# Commands are shared by using modules or snap-ins.

# Modules
# A module is a package of commands and other items that you can use in Windows PowerShell.
# After you run the setup program or save the module to disk, you can import the module into
# your Windows PowerShell session and use the commands and items. You can also use modules
# to organize the providers, functions, aliases, and other commands that you create,
# and share them with others.

Get-Module -ListAvailable
Import-Module Dirsync
Import-Module ActiveDirectory 

# If a module is not present, you will need to install : https://docs.microsoft.com/en-us/powershell/module/powershellget/install-module?view=powershell-7.2
Install-Module -Name O365CentralizedAddInDeployment

Import-Module -Name O365CentralizedAddInDeployment

# After installing and importing the module, we can start using functions from that module
Connect-OrganizationAddInService

Get-OrganizationAddIn

foreach($G in (Get-organizationAddIn)){Get-OrganizationAddIn -ProductId $G.ProductId | Format-List}

# show cmdlets from Dirsync module
Get-Command -module Dirsync
# To retrieve commands from temporary modules downloaded on the fly by a connection script, such as EXO Powershell module, run:
Get-Command -module "tmp*"
Get-Command -module tmp_nqerwha0.c2h

# modules locations
$env:PSModulePath
Get-Content env:\PSModulePath

# Snap-ins
# A Windows PowerShell snap-in (PSSnapin) is a program written in a .NET Framework language 
# that is compiled into a dynamic link library (.dll) that implements
# cmdlets and providers. When you receive a snap-in, you need to install it, and then you
# can add the cmdlets and providers in the snap-in to your Windows PowerShell session.

Get-PSSnapin
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 

# To connecto to EXO we don't need any module; We are connecting using a PS Session (so the module will be automatically downloaded when importing PS Session)
# To connect to MSOLService : https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-msonlinev1?view=azureadps-1.0

#endregion 

#region Execution Policy
<#
The execution policy is part of the security strategy of Windows PowerShell. 
It determines whether you can load configuration files (including your Windows PowerShell profile) 
and run scripts, and it determines which scripts, if any, must be digitally signed before they will  run.

- Restricted: Does not load configuration files or run scripts. "Restricted" is the default execution policy.
- AllSigned: Requires that all scripts and configuration files be signed by a trusted publisher, including scripts that you write on the local computer.
- RemoteSigned: Requires that all scripts and configuration files downloaded from the Internet be signed by a trusted publisher.
- Unrestricted: Loads all configuration files and runs all scripts. If you run an unsigned script that was downloaded from the Internet, you are prompted for permission before it runs.
- Bypass: Nothing is blocked and there are no warnings or prompts.
- Undefined: Removes the currently assigned execution policy from the current scope. This parameter will not remove an execution policy that is set in a Group Policy scope.
#>

#For connecting to Exchange Online
Set-ExecutionPolicy RemoteSigned
#endregion

#region Variable
#TODO - continue from here
# Discuss RunspaceID
# $var - the content of the variable "x"

# ${Any name of the variable between curly brackets}
# When variable names contain - or we need to use curly braces to use the variable name
${variable-name} = "string"

dir variable: #System path where we can explore all current variables in the current powershell session
Get-Variable

# Optional
# Single quotes vs double quotes (and back tick :) )
$var = "Victor"
$v1 = "My name is $var"
$v1
$v2 = 'My name is $var'
<#
using single quotes forces Powershell to use a literal string, not acting on special characters inside the string
#>
$v2
$v3 = "My name is `$var=$var" 
<#
# to display special characters from Powershell scripting language as text, we need to escape the special character by using escape character 
#>
$v3

$mbx = Get-mailbox admin
$mbx.name # displaying the value of a property inside an object returned by the cmdlet - objects to be discussed later
$v4 = " My mailbox is $mbx"
$v4
$v4 = " My mailbox is $mbx.name"
$v4

$v4 = " My mailbox is $($mbx.name)"
$v4

$v5 =  " My mailbox is " + $mbx.name 
$v5 

$var | Get-Member # To view all properties and methods associated to the object stored in your variable


<# Variable Data Types
The most common DataTypes used in PowerShell are listed below.

[string]    Fixed-length string of Unicode characters
[char]      A Unicode 16-bit character
[byte]      An 8-bit unsigned character

[int]       32-bit signed integer
[long]      64-bit signed integer
[bool]      Boolean True/False value

[decimal]   A 128-bit decimal value
[single]    Single-precision 32-bit floating point number
[double]    Double-precision 64-bit floating point number
[DateTime]  Date and Time

[xml]       Xml object
[array]     An array of values
[hashtable] Hashtable object

for an extensive list of Types in Powershell, you can review : https://docs.microsoft.com/en-us/powershell/scripting/lang-spec/chapter-04?view=powershell-7.2

# To force a variable type (always provide variable type in script)
[int]$number = 5
# If we do not specifically set the variable type, Powershell will perform an automatic type casting (selects the type it believes you will need, which can be wrong)
#>

#endregion

#TODO - continue review from here
#region Operators (help about_Comparison_Operators)
<#
Windows PowerShell includes the following comparison operators:
  -eq
  -ne
  -gt
  -ge
  -lt
  -le
  -Like (used for wildcard comparison; used only for string comparison)
  -NotLike
  -Match
  -NotMatch
  -Contains (result: true/false; will tell if one collection of objects contains an object)
  -NotContains
  -In (result: true/false; will tell if one object is included in a collection of objects)
  -NotIn
  -Replace

Other operators:
- is
- as
- Replace
- Join
- Split ","


!!! The match operators search only in strings. They cannot search in arrays.

By default, all comparison operators are case-insensitive. To make a comparison operator case-sensitive, precede the operator name with a "c".
For example, the case-sensitive version of "-eq" is "-ceq".
To make the case-insensitivity explicit, precede the operator with an "i". For example, the explicitly case-insensitive version of "-eq" is "ieq".
#>
#endregion

#region Objects (help about_Objects)
<#
Every action you take in Windows PowerShell occurs within the context of objects.
As data moves from one command to the next, it moves as one or more identifiable objects. 
An object, then, is a collection of data that represents an item.

An object is made up of three types of data:
- the objects type
- its methods
- its properties.
More objects are a Collection (the result of the bellow command).
#>
Get-Mailbox admin | Get-Member

# Get-Member (gm alias)
# Any command that produces output on the screen can be piped to Get-Member in order to see the events, alias properties, methods, properties and note properties
cls
#endregion

#region Pipeline
# The output of one command is used as the input for another command
# If the parameters that we need from the left side is not identical with a parameters on the right side
# we need to create a custom one based the information we have on the left side :) 
Get-Mailbox -PublicFolder |Get-MailboxStatistics

$mailboxes = New-Object Object | Select-Object -Property Name,EmailAddress
$mailboxes.Name = "admin"
$mailboxes.EmailAddress = "admin@vilega.onmicrosoft.com"

#$mailboxes = Import-Csv "test.csv"
$mailboxes | Get-Mailbox
Help Get-Mailbox -Full

$mailboxes | Select-Object @{Name='Identity'; expression = {$_.Name}} | Get-Mailbox
#endregion

#region Select-Object
# From an object we can keep only what properties we need
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress | Get-Member

#What would be the type of this object?
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress,ThrottlingPolicy  |gm

# You can use switches:
#- Last 10
#- First 5

# aproperty can be expanded:
Get-Mailbox admin |Select-Object -ExpandProperty EmailAddresses
$FormatEnumerationLimit = -1
Get-Mailbox admin | fl EmailAddresses

#endregion

#region Sorting: Sort-Object
# You can sort an object on a property. The default sorting is ascending.
# If you need to sort descending you need to specify the switch -Descending
Get-Mailbox |Select-Object Name, PrimarySmtpAddress |Sort-Object Name -Descending
#endregion

#region Formatting & Exporting
<#
Formating
The results can be piped to:
Format-List (alias fl)
Format-Table -Wrap -AutoSize (alias ft -w -A)
Format-Wide (alias fw)

In case no special formatting was created for the command if we have 4 or less properties
the "Format-Table" will be chosen and if there are 5 or more properties the "Format-List" will be used.

!!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.

# Exporting commands results (after |)
get-mailbox admin |fl > test.txt -> write
# get-mailbox admin |fl >> test.txt -> append
Export-csv -Path c:\output1.csv
Export-clixml -Path C:\output2.xml
Import-clixml C:\output2.xml
Out-File -FilePath C:\output3.txt
Out-GridView
ConvertTo-Html | Out-File output4.html
ConvertTo-Csv
ConvertFrom-Json
ConvertTo-Json 
(Convert verb will let the output in the Shell)
(Export = Convert + Out)
#>

Get-Mailbox | Out-GridView
$mbx = Get-Mailbox -SoftDeletedMailbox | Out-GridView -PassThru

#endregion

#region Custom Properties
#@{Name='<Property Name>'; Expression= {<ExpressionValue>}}
Get-Mailbox admin | Select-Object name, PrimarySmtpAddress, @{Name='Mailbox Creation Year'; Expression= {$_.WhenCreated.Year}} 
#endregion

#region Loops

# Conditional Logic - (if, elseif, elseif, switch)
if (condition) {code block}
elseif (condition) {code block}
else  {code block}

Switch (<test-value>)
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
    1 {Write-Host "1"; break}
    2 {Write-Host "2"; break}
    3 {Write-Host "3"; break}
    default {Write-Host "default"}
}
# !!! The "Default" keyword specifies a condition that is evaluated only when no other conditions match the value.
# !!! Any Break statements apply to the collection, not to each value.

# Where-Object
Get-Mailbox | Where-Object {$_.PrimarySmtpAddress -eq "admin@vilega.onmicrosoft.com"}
Get-Mailbox | ? PrimarySmtpAddress -eq "admin@vilega.onmicrosoft.com"

Get-Mailbox |  ? AuditEnabled -eq "False" | set-mailbox -AuditEnabled $true -confirm:$False
Get-Mailbox |  ? LitigationHoldEnabled -eq "True"



<# Conditional Logic - Loops
do while - Script block executes as long as condition value = True.
do while
do {code block}
while (condition)

while - Same as “do while” but the condition is evaluated first
while
while (condition) {code block}

do until – Script block executes until the condition value = True.
do {code block}
until (condition)


for – Script block executes a specified number of times.
for (initialization; condition; repeat/increment/decrement)
{code block}



foreach - Executes script block for each item in a collection or array.
foreach ($<item> in $<collection>)
{code block}



ForEach-Object - Executes script block for each item in a collection or array but is used after pipeline ($_)
$<collection> |  ForEach-Object { code block}
#>

for(i=0; i<10;i++){}
$i=0; for(;;){Write-Host $i; $i++; if($i -eq 10){break;}} #simulates a while(true) loop with a condition to break out of the loop

foreach ($mbx in $mbxs){$mbx}

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

# from external, you'll get the code
$LASTEXITCODE

# from PS you'll get true/false depending on the result of last cmdlet
$?

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

$ErrorActionPreference = "Continue"
$ErrorActionPreference = 'Stop'
try { get-mailbox adfad}
catch{ Write-Host "fadfafad"}
finally {}

#endregion

#region to add a new value, keeping old values on a property (to add/remove email address, domains, IPs, etc...)
Get-Mailbox cloud2 |fl *email*
# To add v1
$mbx =  (Get-Mailbox cloud2).EmailAddresses
$mbx.Add("smtp:cloud2_1@vilega.onmicrosoft.com")
$mbx
Set-Mailbox cloud2 -EmailAddresses $mbx
Get-Mailbox cloud2 |fl *email*

# To remove v1
$mbx =  (Get-Mailbox cloud2).EmailAddresses
$mbx
$mbx.Remove("smtp:cloud2_1@vilega.onmicrosoft.com")
$mbx
Set-Mailbox cloud2 -EmailAddresses $mbx
Get-Mailbox cloud2 |fl *email*

# To add / remove v2
Get-Mailbox cloud2 |fl *email*
Set-Mailbox cloud2 -EmailAddresses @{add="smtp:cloud2_1@vilega.onmicrosoft.com"}
Get-Mailbox cloud2 |fl *email*
Set-Mailbox cloud2 -EmailAddresses @{remove="smtp:cloud2_1@vilega.onmicrosoft.com"}
Get-Mailbox cloud2 |fl *email*

#endregion

<# !!! For the second presenation:
- need to review for, foreach, foreach-object
- need to review try,catch,finally

#>

#region Calculate Hash
# Computes the hash value for a file by using a specified hash algorithm.
# Get-FileHash [-Path] <String[]> [-Algorithm <String> {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160} ] [ <CommonParameters>]

 Get-FileHash C:\Users\vilega\Desktop\test20mb.file -Algorithm SHA1 | Format-List

#endregion

#region Send-MailMessage
# !!! you need to have port 25 opened !!!
# [System.Net.ServicePointManager]::SecurityProtocol = 'Tls,TLS11,TLS12'

$time = Get-Date -Format yyyyMMdd_hhmmss
$creds = Get-Credential
Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -From admin@vilega05.onmicrosoft.com -To admin2@vilega05.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential $creds -Attachments "C:\Users\vilega\Desktop\test.com"
Send-MailMessage -SmtpServer vilega05.mail.protection.outlook.com -From admin@vilega05.onmicrosoft.com -To admin2@vilega05.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential -Attachments "C:\Users\vilega\Desktop\test.com"


#endregion

#region Out-GridView -PassThru (admin interaction for the output)
$mbx = Get-Mailbox |Out-GridView -PassThru
$mbx
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

#region XML - Why?
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

#region Pass the value to the console but also re-use it
Get-mailbox admin2 |select * |Tee-Object -FilePath C:\Users\vilega\OneDrive\Powershell\Trainings\Me\Outputs\GetMailbox.txt 
Invoke-Item C:\Users\vilega\OneDrive\Powershell\Trainings\Me\Outputs\GetMailbox.txt

#endregion

#region Function
Function Test
{
[CmdletBinding(SupportsShouldProcess)]
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

#region Array/HashTable
# @() - Array
# @{} - HashTable
# $hashtable = @{name="Victor";age="36"}

$collection= @()
$object1 = New-Object PSObject
$object2 = New-Object PSObject
$object1 |Add-Member -Name Nume -Value "Victor" -MemberType NoteProperty
$object2 |Add-Member -Name Nume -Value "Andrei" -MemberType NoteProperty 
$collection+= $object1
$collection+= $object2
$collection | foreach{Write-host($($_.Nume))}

#endregion

#region Split

$a = "this is a sample text"
$a.Split(" ")[0]

#endregion 

#region Remember
<#
- Everything is an object
- Check if Customer is having at least minimum required version of PowerShell (example: 3)
- Start-Transcript
- history
- Run "$FormatEnumerationLimit = -1" before collecting information
- If you check all objects please user "-Resultsize Unlimited", or "-All" depends on the cmdlet used
- Always filter on the left side and format only on the right side
- Always run the cmdlet on your side before sending them to customers
- If more cmdlets run to change something, always use Micorosoft disclaimer
- Run the cmdlets first with -WhatIf
- The most important cmdlets are:
    Get-Help
    Get-Command
    Get-Member
- When you have a complex situation, break it in small pieces that can be manageable  and tested
- !!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.
#>


#endregion


#Dif between contains, match,like

"10-20" -Contains "-"
"10-20".Contains("-")
"10-20" |gm

"10-20" -split "" -contains "-"

@("10","20","30") -contains "20"


$AllMbx = get-mailbox 

$AllMbx.GetType()
$mbx= get-mailbox test21
$AllMbx.Contains('test21')
([System.Collections.ArrayList]$AllMbx).Contains($mbx)
$AllMbx.Contains($mbx)

$mbx=$AllMbx[0]
$AllMbx.Contains($mbx)
Start-Process "http://www.computerperformance.co.uk/powershell/powershell_conditional_operators.htm"

get-help System.Collections.ArrayList
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
foreach ($mb in get-mailbox ){$_.alias}
}

(Measure-Command $block1).TotalMilliseconds
(Measure-Command $block2).TotalMilliseconds

foreach ($character in [char[]]”Poshoholic”) { if (@(‘a’,’e’,’i’,’o’,’u’) -contains $character ) { continue } $character }
[char[]]”Poshoholic” | foreach { if (@(‘a’,’e’,’i’,’o’,’u’) -contains $_ ) { continue } $_ }

$var1 = @("1","2")
$var2 = "1","2"
$var1.GetType()
$var2.GetType()


#Begin
$path=[Environment]::GetFolderPath("Desktop")
#$path = "c:\temp" 
$timestamp = Get-Date -format yyMMdd_hhmmss
Start-Transcript -Path "$Path\Transcript_$timestamp.txt" -Force

Stop-transcript
#End




get-mailbox admin |fl *copy*

[string]::IsNullOrEmpty


Get-OrganizationConfig |fl *block*
help Set-OrganizationConfig -Parameter IPListBlocked

Get-ClientAccessRule


$calendars = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Get-MailboxFolderStatistics | ? {$_.FolderType -eq "Calendar"} | select @{n="Identity"; e={$_.Identity.Replace("\",":\")}}
$calendars | % {if ((Get-MailboxFolderPermission -Identity $_.Identity -User Default).AccessRights -ne "Reviewer") {Set-MailboxFolderPermission -Identity $_.Identity -User Default -AccessRights Reviewer}} 

$calendars | % {Set-MailboxFolderPermission -Identity $_.Identity -User Default -AccessRights LimitedDetails} 



[string]$myvar2 = 1
$myvar1 = 1
$myvar1.GetType()
Get-Command -ParameterName UserPrincipalName

@( “machine1” , “machine2” , “machine3”).GetType()