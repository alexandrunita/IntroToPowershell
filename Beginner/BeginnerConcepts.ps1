#region WhatIsPowerShell
Write-Host "Windows PowerShell is an interactive object-oriented command environment with scripting language features that utilizes small programs called cmdlets to simplify configuration, administration, and management of heterogeneous environments in both standalone and networked typologies by utilizing standards-based remoting protocols."
<# Powershell editors:
- PowerShell
- Integrated Scripting Environment (ISE)
- 3rd party editors 
- Windows terminal
- Visual Studio Code
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

# aliases are not supported to use with Invoke-Command, and are not recommended in scripting

#endregion

#region Check Environment Settings

# Check Powershell version
$PSVersionTable
# to check if your PS module has a minimum requirement of PS version, or there may be some know issues with newer versions of PS

# Console Settings
Get-Host
# WWindows console
# EMS, etc: host that loads PS and automatically loading module and connect to workload (right click on shortcut to see what it does at startup)

# Local Machine settings
Get-Culture
# note the culture that PowerShell is using (in this case en-US)

#If you need to run your PowerShell in a lower version, open Powershell with the following command: "powershell.exe -version 2"
#In this case Powershell version used was 2
# older PS verions were based on Windows Management Framework
#endregion

#region HelpTool

# first thing to do on a new machine
# Update Help - lengthy
# Update-Help

# If we want to save Help content to a file, run:
Save-Help

# examples for which is important to have updated helpo repository
help about*
help about_Operators
help about_Aliases
help about_Arithmetic_Operators
# help about* - very useful for learning usage scenario

# We can also update-help from a saved file via save-help if without Internet Connection
Update-Help -SourcePath %%

# Use powershell help first instead of any search engine
#Default command
get-help Get-Mailbox 
#If not connected to Workload Module, help will not retrieve the information
# there are parameter sets which may be excluding one other - SYNTAX section

# To only retrieve help on a specific parameter:
get-help Get-Service -Parameter name

# to see the examples:
get-help Set-Mailbox -examples

# for more information
get-help Set-Mailbox -detailed
get-help Set-Mailbox -full 
# for technical information
get-help Get-Mailbox -Parameter Identity
#useful inscripting for best identifier usage (ex: identifier may not accept wildcard)

# for online help 
get-help Set-Mailbox -online

# to explore all options of Get-Help cmdlet, check the help repository:
Get-Help Get-Help
# Help structured on topics (like user manual for powershell). Recommend to use it !!!

# We can use alias "help" for Get-Help cmdlet
help about


#endregion HelpTool

#region Get-Command
# List of the commands available. You can use wildcards to find it. (the results will be:  cmdlets, functions, workflows, and aliases )
# Alias for "Get-Command" is "gcm"
# Details retrived from local module/dll
# Will not retrieve information if module is not loaded, except for installed module ExportedCommands

Get-Command Get*MSOL*
Get-Command *Calendar*
Get-Command *set*Calendar*

Get-Command -ParameterName max*
Get-Command -ParameterName UserPrincipalName
Get-Command -ParameterName UserPri*

# Get-Member (gm alias)
#Any command that produces output on the screen is either an object or a collection of objects, and can be piped to Get-Member in order to see the events, alias properties, methods, properties and note properties

# to check object type of each entry
Get-Mailbox | Get-Member

# to check the type of output object/collection
(Get-Service).GetType()

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
# id there is no confusion, PS will implicitly use the proper one (Id, Ide, Iden)

<#
If the cmdlet is making changes you can use the following parameters:
- WhatIf (to see what will be changed)
-Confirm:$False (in case a permission need to be asked, it will automatically proceed)
- note -Confirm:$False is a switch and it requires ":". If it hase defalut value you can simply invoke it, example: -HasErrorsOnly for Get-MsolUser
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
# recall what you've ran 
# see all cmdlets ran if you forgot to start transcript - export for documenting

Invoke-History 
# can re-run a specific cmdlet from history - position no. based
r 34
get-alias r
#endregion

#region Snapin and Modules
# Commands are shared by using modules or snap-ins.
    # snap-ins - old way of grouping cmdlets. It loads all cmdlets, but if you don't have RBAC permissions on some, they won't work.
# Modules
# A module is a package of commands and other items that you can use in a Windows PowerShell session.
# After you run the setup program or save the module to disk, you can import the module into
# your Windows PowerShell session and use the commands and items. You can also use modules
# to organize the providers, functions, aliases, and other commands that you create,
# and share them with others.
    # In EXO PS session, you'll only have available the cmdlets for which you have permissions via RBAC
    # Other workloads can include all cmdlets in module

Get-Module
Get-Module -ListAvailable
    Import-Module Dirsync
    Import-Module ActiveDirectory 

# If a module is not present, you will need to install : https://docs.microsoft.com/en-us/powershell/module/powershellget/install-module?view=powershell-7.2
Install-Module -Name O365CentralizedAddInDeployment

Import-Module -Name O365CentralizedAddInDeployment

# -Force parameter - can be used with install/import module; if module is already loaded in the session

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
# Each cmdlet run with |fl in modules used via Powershell Remoting will return a RunspaceID

# This RunspaceID uniquely identifies the Powershell session used to run the cmdlets
# The RunspaceID has no relation to the object/configuration retrieved by the cmdlet

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
- the objects type (string/integer/float number/array/etc..)
- its methods (aka functions/actions)
- its properties. (aka attributes)
More objects are a Collection (the result of the bellow command).
#>
Get-Mailbox admin | Get-Member

# Get-Member (gm alias)
# Any command that produces output on the screen can be piped to Get-Member in order to see the events, methods, properties
#endregion Objects

#region Custom Properties
#@{Name='<Property Name>'; Expression= {<ExpressionValue>}}
Get-Mailbox admin | Select-Object name, PrimarySmtpAddress, @{Name='Mailbox Creation Year'; Expression= {$_.WhenCreated.Year}} 
#endregion Custom Properties

#region Pipeline
# The output of one command is used as the input for another command
# If the parameters that we need from the left side are not identical with a parameters on the right side
# we need to create a custom one based the information we have on the left side

# For example, we cannot get the MSOLGroup as ObjectID Parameter is not found on the output of the cmdlet on the left:
Get-DistributionGroup ADList@liviunita.onmicrosoft.com | Get-MsolGroup

# To make this, work we can do something like:

$mailboxes = New-Object Object | Select-Object -Property Name,EmailAddress
$mailboxes.Name = "liviu"
$mailboxes.EmailAddress = "liviu@liviunita.onmicrosoft.com"

# alternatively, we can import from a CSV $mailboxes = Import-Csv "test.csv"
$mailboxes | Get-Mailbox
Help Get-Mailbox -Full

$mailboxes | Select-Object @{Name='Identity'; expression = {$_.EmailAddress}} | Get-Mailbox

# Going back to example of the Distribution, we can fix the pipeline input matching for MsolGroup like this:
Get-DistributionGroup ADList@liviunita.onmicrosoft.com | Select-Object @{Name='ObjectId'; expression = {$_.ExternalDirectoryObjectId}}| Get-MsolGroup

#endregion

#region Select-Object
# From an object we can keep only what properties we need
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress | Get-Member

#What would be the type of this object?
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress,ThrottlingPolicy  |gm

# You can use switches:
#- Last 10
#- First 5
Get-Mailbox -ResultSize Unlimited | select -last 10

# a property can be expanded:
Get-Mailbox admin |Select-Object -ExpandProperty EmailAddresses
(Get-Mailbox admin).EmailAddresses
$FormatEnumerationLimit = -1
Get-Mailbox admin | fl EmailAddresses
#endregion

Get-Mailbox admin | Select-Object -ExcludeProperty UserCertificate 
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

In case no special formatting was created for the command if we have 4 or less properties
the "Format-Table" will be chosen and if there are 5 or more properties the "Format-List" will be used.

!!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.

# Exporting commands results (after |)
get-mailbox admin |fl > test.txt -> write
# get-mailbox admin |fl >> test.txt -> append
Export-csv -Path c:\output1.csv -NoTypeInformation

Export-clixml -Path C:\output2.xml -Depth 5 
#default depth is 2, which may not always be sufficient to expand all objects, for example:
$msolUser = Read-Host
(Get-MsolUser -UserPrincipalName $msolUser).licenses[0].servicestatus

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

Get-Mailbox | select -First 10 | Out-GridView
$mbx = Get-Mailbox -SoftDeletedMailbox | Out-GridView -OutputMode single
$mbxs = Get-Mailbox -SoftDeletedMailbox | Out-GridView -PassThru
#endregion


#region Loops

# Conditional Logic - (if, elseif, else, switch)
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

# TODO - check if we need to add if/elseif/else examples

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

# foreach - keyword
foreach ($mbx in $mbxs){$mbx}

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
try { get-mailbox -Erroraction Stop adfad}
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



#region Array/HashTable
# @() - Array - data collection, indexable, immutable
# @{} - HashTable - key value collection (index - unsorted key value pair)
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
foreach ($mb in get-CASmailbox ){$_.alias}
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