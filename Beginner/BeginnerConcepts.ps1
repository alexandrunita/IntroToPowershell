#region WhatIsPowerShell
Write-Host "PowerShell is a cross-platform task automation solution made up of a command-line shell, a scripting language, and a configuration management framework. PowerShell runs on Windows, Linux, and macOS."

# PowerShell is built on the .NET Common Language Runtime (CLR). All inputs and outputs are .NET objects. 
#For full description: https://docs.microsoft.com/en-us/powershell/scripting/overview?view=powershell-7.2

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

# aliases are not supported to use with Invoke-Command, and are not recommended in scripting !!!

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
# help about* - very useful for learning usage scenario !!!

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
Get-Service | Get-Member

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
Get-History
History
history | select -First 1 | fl *
History | select CommandLine
history | select -last 10
# recall what you've ran 
# see all cmdlets ran if you forgot to start transcript - export for documenting
Get-History | Out-File -FilePath c:\PS1\session_history.txt

Invoke-History 
# can re-run a specific cmdlet from history - position no. based
r 34
get-alias r

# clear history
Clear-History
Clear-History -Count 5 -Newest
(Get-PSReadlineOption).HistorySavePath
Get-Content -Path (Get-PSReadlineOption).HistorySavePath
Clear-Content -Path (Get-PSReadlineOption).HistorySavePath

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
    Import-Module Msonline
    Import-Module ActiveDirectory 

# If a module is not present, you will need to install : https://docs.microsoft.com/en-us/powershell/module/powershellget/install-module?view=powershell-7.2

# -Force parameter - can be used with install/import module; if module is already loaded in the session

# After installing and importing the module, we can start using cmdlets from that module

# example2: show cmdlets from Msonline module
Get-Command -module Msonline
# To retrieve commands from temporary modules downloaded on the fly by a connection script, such as EXO Powershell module, run:
Get-Command -module "tmp*"
Get-Module "tmp*"
Get-Command -module tmp_5wmmdo5h.rt0

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

# Single quotes vs double quotes (and back tick :) )
$var = "Alex"
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

For additional details on comparison operators, you can review: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comparison_operators?view=powershell-7.2
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
More objects constitute a Collection.
#>

Get-Mailbox admin | Get-Member

# Get-Member (gm alias)
# Any command that produces output on the screen can be piped to Get-Member in order to see the events, methods, properties
#endregion Objects

#region Custom Properties
#@{Name='<Property Name>'; Expression= {<ExpressionValue>}}

(Get-Mailbox admin).WhenCreated | Get-Member
(Get-Mailbox admin).WhenCreated.Year
(Get-Mailbox admin).WhenCreated.ToUniversalTime()

Get-Mailbox admin | Select-Object name, PrimarySmtpAddress, @{Name='Mailbox Creation Year'; Expression= {$_.WhenCreated.Year}} 
#endregion Custom Properties

#region Pipeline
# The output of one command is used as the input for another command
# If the parameters that we need from the left side are not identical with a parameters on the right side
# we need to create a custom one based the information we have on the left side

# For example, we cannot get the MSOLGroup as ObjectID Parameter is not found on the output of the cmdlet on the left:
Get-DistributionGroup testalert1@axul.onmicrosoft.com | Get-MsolGroup

Get-Command Get-MsolGroup -Syntax  

# Before making the above command work, let's create a custom object, define its properties and add values to them:

$mailbox = New-Object Object | Select-Object -Property Name,PrimarySmtpAddress
$mailbox.Name = "testalert1"
$mailbox.PrimarySmtpAddress = "testalert1@axul.onmicrosoft.com"
$mailbox

# If we want to create a custom object based on the output of a cmdlet with specific property names:

$mailbox2 = Get-Mailbox admin@axul.onmicrosoft.com | Select-Object @{Name='Label1'; expression = {$_.PrimarySmtpAddress}}, @{Name='Label2'; expression = {$_.ExternalDirectoryObjectId}}
$mailbox2

# Going back to example of the Distribution, we can fix the pipeline input matching for MsolGroup like this:
Get-DistributionGroup testalert1@axul.onmicrosoft.com | Select-Object @{Name='ObjectId'; expression = {$_.ExternalDirectoryObjectId}}| Get-MsolGroup

#endregion

#region Select-Object
# From an object we can keep only what properties we need
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress | Get-Member

#What would be the type of this object?
Get-Mailbox admin |Select-Object Name, PrimarySmtpAddress,ThrottlingPolicy  |gm

# You can use switches:
#- Last 10
#- First 5
Get-Mailbox -ResultSize Unlimited | select -last 5

# a property can be expanded:
Get-Mailbox admin | Select-Object -ExpandProperty EmailAddresses
(Get-Mailbox admin).EmailAddresses
$FormatEnumerationLimit = -1
Get-Mailbox admin | fl EmailAddresses

Get-Mailbox admin | Select-Object -ExcludeProperty UserCertificate 
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

In case no special formatting was created for the command, if we have 4 or less properties
the "Format-Table" will be automatically chosen and if there are 5 or more properties the "Format-List" will be used.

!!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.

# Exporting commands results (after |)
get-mailbox admin |fl > test.txt -> write
# get-mailbox admin |fl >> test.txt -> append

get-mailbox admin | Export-csv -Path c:\output1.csv -NoTypeInformation

get-mailbox admin | Export-clixml -Path C:\output2.xml -Depth 5 

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

# useful to export as XML for offline troubleshooting:

get-mailbox -ResultSize unlimited | Export-clixml -Path C:\PS1\all_mailboxes.xml

$mbxs = Import-clixml -Path C:\PS1\all_mailboxes.xml

$mbxs | Where-Object {$_.RecipientTypeDetails -like 'SharedMailbox'} | Select-Object alias, PrimarySmtpAddress, UserPrincipalName | Sort-Object UserPrincipalName -Descending | ft
#endregion

# Where-Object
Get-Mailbox | Where-Object {$_.PrimarySmtpAddress -eq "admin@axul.onmicrosoft.com"}
Get-Mailbox | ? PrimarySmtpAddress -eq "admin@axul.onmicrosoft.com"

Get-Mailbox |  ? AuditEnabled -eq "False" | set-mailbox -AuditEnabled $true -confirm:$False
Get-Mailbox |  ? LitigationHoldEnabled -eq "True"


