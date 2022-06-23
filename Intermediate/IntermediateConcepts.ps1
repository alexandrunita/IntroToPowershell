
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
- When you have a complex situation, break it in small pieces that can be manageable and tested (user write-host to view the values first)
- !!! After an output was formatted you cannot export to CSV, XML !!! You can only out to host, file (txt), printer, string.
#>

#endregion Remember

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


# foreach alias for Foreach-Object cmdlet
$mbxs = Get-Mailbox
$mbxs | foreach {$_}
get-mailbox | foreach {$_}
$mbxs | ForEach-Object {$_.alias}

# example for setting calendar folder permissions for user default to Reviwer on all mailboxes of type UserMailbox, using foreach (% is foreach PS alias):
$calendars = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Get-MailboxFolderStatistics | ? {$_.FolderType -eq "Calendar"} | select @{n="Identity"; e={$_.Identity.Replace("\",":\")}}
$calendars | % {if ((Get-MailboxFolderPermission -Identity $_.Identity -User Default).AccessRights -ne "Reviewer") {Set-MailboxFolderPermission -Identity $_.Identity -User Default -AccessRights Reviewer}} 


For($i=0;$i -le 4;$i++) {
    write-host($i);
}


for ($i=0;$i -lt ($mbxs.length); $i++)
{
    $mbxs[$i].alias
}

$mbxs.length
$mbxs.count # count is alias for length, for more details check the help about arrays: help about*array*

$a = @()  # new array
$a.GetType().getproperties() | ?{$_.Name -match "length"}

$a.get_Length()


#endregion Loops

#region Errors

# Muliple ways to do it, for more details: https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.2
# also check: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_try_catch_finally?view=powershell-7.2
# To keep it simple you can use bellow command to see last error details


$error # all errors from the current session - array
$error[0] # last error in $error array
$error[0].Exception | fl * -Force

#additional last error details:
$e = $error[0];
$e.Exception | fl * -f
$e.Exception.SerializedRemoteException | fl * -f

# from PS you'll get true/false depending on the result of last cmdlet
$?

# try / catch

try { get-mailbox admin -Erroraction Stop}
catch{ Write-Host "Something went wrong..."}
finally {write-host "Regardless..."}


$ErrorActionPreference = "Continue"
$ErrorActionPreference = 'Stop'

get-mailbox joker
get-mailbox admin
Get-Recipient joker

# Example 1 (catch all):
try {get-mailbox joker -Erroraction Stop}
catch{ 
    If ($error[0].Exception -match "couldn't be found on") {
        $error.Clear()
        Get-Recipient joker -ErrorAction continue
        if ($error[0].Exception -match "couldn't be found on") { write-host "Could not find a mailbox or recipient"}
        
        if(!$error) {Write-Host "Recipient found"}
    }
}   
finally {write-host "Check finished"}

# Inside the catch block, there's an automatic variable ($PSItem or $_) of type ErrorRecord that contains the details about the exception.
# Example 2 (multiple catch needs to run as script, not as chain of cmdlets - workaround if () {} before try/multiple catch)

if ($true) {
    try {get-mailbox joker -Erroraction Stop
        get-joker    
    }
    catch [System.Management.Automation.RemoteException] { 
        If ($error[0].Exception -match "couldn't be found on") {
            Get-Recipient joker -ErrorAction silentlycontinue
            if ($PSItem.Exception -match "couldn't be found on") { write-host "Could not find a mailbox or recipient"}
            
            if(!$error) {Write-Host "Recipient found"}
        }
    }
    catch {"An error occurred that could not be resolved."}   
    finally {write-host "Check finished"}
}

# ErrorAction silentlycontinue does not increment $error variable, but if you are inside catch, you'll see the error in $PSItem.Exception

$error.Clear()
Get-Recipient joker -ErrorAction silentlycontinue
$error[0]

#endregion Errors


#region to add a new value, keeping old values on a property (to add/remove email address, domains, IPs, etc...)

Get-Mailbox user8 |fl *emaila*

# To add v1
$mbx =  (Get-Mailbox user8).EmailAddresses
    $mbx.GetType() # to check object type - array list
    $mbx.GetType() | Get-Member # to check all methods available
    $mbx.GetType().GetMethods().name
    $mbx.GetType().GetMethod("Add")
    help ArrayList # see below comment
    help about_methods
$mbx.Add("smtp:user8abc@axul.onmicrosoft.com")
Set-Mailbox user8 -EmailAddresses $mbx
Get-Mailbox user8 |fl *emaila*

# Array lists are actually .NET collection objects and are just used by PS (not built into PS). If we want to read about such .net objects (and their methods and so on) we need to read the .NET documentation, example: https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=net-6.0

# To remove v1
$mbx =  (Get-Mailbox user8).EmailAddresses
$mbx
$mbx.Remove("smtp:user8abc@axul.onmicrosoft.com")
$mbx
Set-Mailbox user8 -EmailAddresses $mbx
Get-Mailbox user8 |fl *emaila*

# To add / remove v2
Get-Mailbox user8 |fl *emaila*
Set-Mailbox user8 -EmailAddresses @{add="smtp:user8abc@axul.onmicrosoft.com"} # @ creates on-the-spot an object of type hashtable (key-value pair), and EXO recognizes and instead of overrite will the specified methor (add)
Get-Mailbox user8 |fl *emaila*
Set-Mailbox user8 -EmailAddresses @{remove="smtp:user8abc@axul.onmicrosoft.com"}
Get-Mailbox user8 |fl *emaila*

#endregion

#region Split
# dotnet embedded function to split string

$a = "this is a sample text"
$a.Split(" ")[0]

"10-20" -Contains "-"
"10-20".Contains("-")
"10-20" |gm

"10-20" -split "" -contains "-"

@("10","20","30") -contains "20"


$a | Get-Member
$a.ToUpper()

# to check all methods for string object, check .NET article: https://docs.microsoft.com/en-us/dotnet/api/system.string?view=net-6.0

#endregion Split

#region Array/HashTable
# @() - Array - data collection, indexable, immutable

# Array example: 

$collection= @()
$object1 = New-Object PSObject
$object2 = New-Object PSObject
$object1 | Add-Member -Name Name -Value "Liviu" -MemberType NoteProperty
$object2 | Add-Member -Name Name -Value "Laurentiu" -MemberType NoteProperty 
$collection+= $object1
$collection+= $object2
$collection | foreach{Write-host($($_.Name))}

$collection.GetType()

# above example has disadvantage with high number of entries - for each new entry added, a new array is created, populated with all members from initial one plus the new member, then delete the old array

# ArrayList example

[System.Collections.ArrayList]$collectionlist= @()
$object1 = New-Object PSObject
$object2 = New-Object PSObject
$object1 | Add-Member -Name Name -Value "Liviu" -MemberType NoteProperty
$object2 | Add-Member -Name Name -Value "Laurentiu" -MemberType NoteProperty 
$collectionlist.Add($object1)
$collectionlist.Add($object2)
$collectionlist | foreach{Write-host($($_.Name))}

$collectionlist.GetType()

# ArrayList will initially occupy more memory than Array, but adding elements does not incur high demand on resources 

# performance test - array vs arraylist:

$blocka = {
    $collection= @()

    for($i=0; $i -lt 1000;$i++){
        
        write-host $i
        $object1 = New-Object PSObject
        $object2 = New-Object PSObject
        $object1 | Add-Member -Name Name -Value "Liviu$i" -MemberType NoteProperty
        $object2 | Add-Member -Name Name -Value "Laurentiu$i" -MemberType NoteProperty 
        $collection+= $object1
        $collection+= $object2
    }
}

(Measure-Command $blocka).TotalMilliseconds

    # $collection | foreach{Write-host($($_.Name))}


$blockal = {
        [System.Collections.ArrayList]$collectionlist= @()
    
    for($i=0; $i -lt 1000;$i++){
            
            write-host $i
            $object1 = New-Object PSObject
            $object2 = New-Object PSObject
            $object1 | Add-Member -Name Name -Value "Liviu$i" -MemberType NoteProperty
            $object2 | Add-Member -Name Name -Value "Laurentiu$i" -MemberType NoteProperty 
            $collectionlist.Add($object1)
            $collectionlist.Add($object2)
    }
}
    
(Measure-Command $blockal).TotalMilliseconds   

# Optimizing PowerShell Scripts
Start-Process "https://blogs.technet.microsoft.com/ashleymcglone/2017/07/12/slow-code-top-5-ways-to-make-your-powershell-scripts-run-faster/"
Start-Process "https://blogs.technet.microsoft.com/heyscriptingguy/2014/05/18/weekend-scripter-powershell-speed-improvement-techniques/"
Start-Process "https://social.technet.microsoft.com/wiki/contents/articles/11311.powershell-optimization-and-performance-testing.aspx"
Start-Process "https://blogs.technet.microsoft.com/heyscriptingguy/2014/05/17/weekend-scripter-best-practices-for-powershell-scripting-in-shared-environment/"
Start-Process "https://blogs.msdn.microsoft.com/powershell/2008/01/28/lightweight-performance-testing-with-powershell/"


# get-help System.Collections.ArrayList

# Hashtable example: !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# @{} - HashTable - key value collection (index - unsorted key value pair)
# $hashtable = @{name="Liviu";age="32"}

# get more info on hashtables
Get-Help about_Hash_Tables 
get-help ConvertFrom-StringData

$A = ConvertFrom-StringData -StringData "Top = Red `n Bottom = Blue"
$A

<#
Name             Value
----             -----
Bottom           Blue
Top              Red
#>

$ageList = @{
    Kevin = 36
    Alex  = 9
}


# add/remove method 1
$key = 'Kevin'
$value = 36
$ageList.add( $key, $value )

$ageList.add( 'Alex', 9 )
$ageList.remove( 'Alex')
$ageList

# add method 2
$ageList = @{}

$key = 'Kevin'
$value = 36
$ageList[$key] = $value
#add or override
$ageList['Alex'] = 9

$environments = @{
    Prod = 'SrvProd05'
    QA   = 'SrvQA02'
    Dev  = 'SrvDev12'
}

$environments[@('QA','DEV')]
$environments[('QA','DEV')]
$environments['QA','DEV']

$environments.values

# loads all keys from beginning
foreach($key in $ageList.keys)
{
    $message = '{0} is {1} years old' -f $key, $ageList[$key]
    Write-Output $message
}

# loads keys one by one, feeing up previous hey from memory - lower on resources
$ageList.GetEnumerator() | ForEach-Object{
    $message = "$($_.key) is $($_.value) years old!"
    Write-Output $message
}

## for more hashtable details, please visit: https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-hashtable?view=powershell-7.2


#endregion


#region create object containing outputs of multiple PS cmdlets (PSCustomObject)

# Example, collecting Get-PublicFolder and Get-MailPublicFolder:

#Create object:

$ipath = "\"
$PFs = Get-PublicFolder $ipath -ResultSize unlimited -Recurse | ?{$_.MailEnabled -eq "True"}
$PFs

$PFInfo = $PFs | foreach {
    [string]$temp = $_.MailRecipientGuid
    New-Object psobject -Property @{
        PFObject = $_
        REC = Get-Recipient $temp
        MEPF = $_ | Get-MailPublicFolder 
    }
}

# Export object: 
$PFInfo | Export-Clixml $env:userprofile\desktop\PFInfoXML.xml -Depth 10

# Load object:

$PFInfoXML = import-clixml $env:userprofile\desktop\PFInfoXML.xml

# Display subset option 1:
Foreach ($item in $PFInfo) # or PFInfoXML
{
    Write-Host "Identity:" $item.PFObject.Identity
    Write-Host "Forwarding:" $item.MEPF.DeliverToMailboxAndForward
    Write-Host "PrimarySmtpAddress:" $item.MEPF.PrimarySmtpAddress
    Write-Host "ContentMailbox:" $item.MEPF.ContentMailbox 
    Write-Host "RecipientTypeDetails:" $item.REC.RecipientTypeDetails
    Write-Host
}

#Display subset option2:

$PFInfoSubset = $PFInfoXML | foreach {
    New-Object psobject -Property @{
        Identity = $_.PFObject.Identity
        PrimarySmtpAddress = $_.MEPF.PrimarySmtpAddress
        Forwarding = $_.MEPF.DeliverToMailboxAndForward
        RecipientTypeDetails = $_.REC.RecipientTypeDetails
        ContentMailbox = $_.MEPF.ContentMailbox 

    }
}

$PFInfoSubset | sort PrimarySmtpAddress | Out-GridView

$PFInfo.GetType()
$PFInfoXML.GetType()
$PFInfoSubset.GetType()

#endregion


# !!! Highly recommended self-study PS resource: https://docs.microsoft.com/en-us/powershell/scripting/how-to-use-docs?view=powershell-7.2


#region Remember

<#
- Don't over complicate things. Keep it simple and use the most straight forward way to accomplish a task. 
- Avoid aliases and positional parameters in any code that you reuse. Format your code for readability. 
- Don't hardcode values; use parameters and variables. Don't write unnecessary code even if it doesn't hurt anything. It adds unnecessary complexity. 
- Attention to detail goes a long way when writing any PowerShell code.
- Always test/visualize first. Use -WhatIf parameter, or use Write-Host to visualize current variable value instead of setting anything with them. When dealing with lists of objects (like users), you can export to csv and visually inspect the list to confirm first.
- For higher complexity issues, where is necessary to colect multiple PS outputs, use PS custom object and export to XML with higher depth (but not higher than needed as it will increase the XML size), then import XML on your side to filter, arrange and so on for analysis troubleshooting.
- Whenever you run PS on customer tenant, even if the cmdlets are only to export some logs in csv or xml file, do it after you start transcript with -IncludeInvocationHeader parameter to record also the cmdlets issued and timestamp in the transcript file. Example:

    #Begin
    $path=[Environment]::GetFolderPath("Desktop")
    #$path = "c:\temp" 
    $timestamp = Get-Date -format yyMMdd_hhmmss
    Start-Transcript -IncludeInvocationHeader -Path "$Path\Transcript_$timestamp.txt" -Force
    # insert here your PS cmdlets
    Stop-transcript
    #End

#>

#endregion Remember

#region functions

# When you have to do the same operation multiple times in a script, using functions is more efficient and clean
# For naming your functions, it is recommended to use approved verbs and to prefix the noun
#    Get-Verb | Sort-Object -Property Verb

# Example of function which returns all cmdlets having specific parameter names:

    function Get-MrParameterCount {
        param (
            [string[]]$ParameterName
        )

        foreach ($Parameter in $ParameterName) {
            $Results = Get-Command -ParameterName $Parameter -ErrorAction SilentlyContinue

            [pscustomobject]@{
                ParameterName = $Parameter
                NumberOfCmdlets = $Results.Count
            }
        }
    }

Get-MrParameterCount -ParameterName emailaddresses, alias, userprincipalname
Get-Command -ParameterName emailaddresses

# for more on functions in PS, please check: https://docs.microsoft.com/en-us/powershell/scripting/learn/ps101/09-functions?view=powershell-7.2

#endregion functions


#region various examples

# example for measuring script block
$block1={
get-mailbox |foreach {$_.alias}
}

$block2={
foreach ($mb in get-CASmailbox ){$_.alias}
}

(Measure-Command $block1).TotalMilliseconds
(Measure-Command $block2).TotalMilliseconds


# Recomendation when taking transcript files: Use the -IncludeInvocationHeader parameter to record also the cmdlets issued and timestamp in the transcript file:

#Begin
$path=[Environment]::GetFolderPath("Desktop")
#$path = "c:\temp" 
$timestamp = Get-Date -format yyMMdd_hhmmss
Start-Transcript -IncludeInvocationHeader -Path "$Path\Transcript_$timestamp.txt" -Force
# insert here your PS cmdlets
Stop-transcript
#End

# Check if a string is null/empty
$string1 = $null
IF ([string]::IsNullOrWhitespace($string1)){'empty'} else {'not empty'}


# Example: it is recommended to declare variable type, otherwise PS will guess
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



# example to continuously check the status of the issued request and exit with a message when request completed

New-MoveRequest user8
Get-MoveRequestStatistics user88 | fl status
    Do {
            $stats = Get-MoveRequestStatistics -Identity "user88"
            Write-Host '*' -NoNewline
            Start-Sleep -Seconds 5
        } While ($stats.Status -ne 'Completed')
        Write-Host "`n> Move request completed"



#Example Pass the value to the console but also re-use it
Get-mailbox admin2 |select * | Tee-Object -FilePath C:\mstemp1\GetMailbox.txt 
Invoke-Item C:\mstemp1\GetMailbox.txt


# example Send-MailMessage
# !!! you need to have port 25 opened !!!
# [System.Net.ServicePointManager]::SecurityProtocol = 'Tls,TLS11,TLS12'

$time = Get-Date -Format yyyyMMdd_hhmmss
$creds = Get-Credential
Send-MailMessage -SmtpServer smtp.office365.com -Port 587 -UseSsl -From admin@axul.onmicrosoft.com -To admin2@axul.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential $creds -Attachments "C:\mstemp1\fstat.txt"
Send-MailMessage -SmtpServer axul.mail.protection.outlook.com -From admin@axul.onmicrosoft.com -To admin2@axul.onmicrosoft.com -Subject "Email_$time" -Body "This is only a test" -Credential -Attachments "C:\mstemp1\fstat.txt"

# example2, for sending multiple emails, for testing purposes, to populate new mailboxes in your test tenant for repro:

[System.Net.ServicePointManager]::SecurityProtocol = 'Tls,TLS11,TLS12'
$cred=Get-Credential
for($i=1; $i -le 200; $i++) {
for($j=1; $j -le 20; $j++) {
Send-MailMessage -UseSSL -Port 587 -Credential $cred -From 'admin <admin@axul.ro>' -To 'stest1 <stest1@axul.ro>', 'stest2 <stest2@axul.ro>' -Subject "Batch$i Item$j - Sending the Attachment" -Body "Forgot to send the attachment. Sending now." -Attachments .\goodfood.jpg -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer 'smtp.office365.com'
}
}


#example for getting MessageTrace details 
Get-MessageTrace |select -Last 1 | Get-MessageTraceDetail |fl

$msgID = (Get-MessageTrace |select -Last 1).MessageId
Get-MessageTrace -MessageId $msgID  | Get-MessageTraceDetail |fl

Get-MessageTrace -MessageId "input here the message id value"  | Get-MessageTraceDetail |fl
Get-MessageTrace | Out-GridView -PassThru |Get-MessageTraceDetail 
Get-MessageTrace | Out-GridView -PassThru |Get-MessageTraceDetail |fl


#region XML - manipulation of XML Reports

#Example: Find MSOL user Error
#Find all users with errors
Get-MsolUser -MaxResults 10000 -HasErrorsOnly

$UPN = "user5@axul.onmicrosoft.com"
(Get-MSOLUser -UserPrincipalName  $UPN).errors.errorDetail.objectErrors.errorRecord.errorDescription 

# check licensing status
$UPN = "admin@axul.onmicrosoft.com"

# First Example - Depth:
$MSOLUserAll = Get-MsolUser -UserPrincipalName $UPN
$MSOLUserAll
$MSOLUserAll.Licenses
$MSOLUserAll.Licenses.ServiceStatus

Get-MsolUser -UserPrincipalName $UPN | Export-Clixml C:\MSTEMP1\MSOLUser.xml -Force
$MSOLUser = Import-Clixml C:\MSTEMP1\MSOLUser.xml
$MSOLUser
$MSOLUser.Licenses
$MSOLUser.Licenses.ServiceStatus

Get-MsolUser -UserPrincipalName $UPN | Export-Clixml C:\MSTEMP1\MSOLUserD.xml -Depth 4 -Force
$MSOLUserD = Import-Clixml C:\MSTEMP1\MSOLUserD.xml
$MSOLUserD
$MSOLUserD.Licenses
$MSOLUserD.Licenses.ServiceStatus


# Second Example -Migration Report
# Explore Migration XML, check why Depth > 2 does not give more info
# Collect from customer
$Mailbox = 'MigratedMailbox'
Get-MoveRequest
#Get-MoveRequestStatistics $Mailbox -IncludeReport -Diagnostic -DiagnosticArgument Verbose | Export-Clixml C:\MSTEMP1\MoveRequestStatistics.xml
Get-MoveRequestStatistics $Mailbox -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" | Export-Clixml C:\MSTEMP1\MoveRequestStatistics_$Mailbox.xml 

# Analyze on the engineer side
$r = Import-Clixml C:\MSTEMP1\MoveRequestStatistics.xml
$r | fl | Out-File C:\MSTEMP1\MoveRequestStatistics.txt -Force; Invoke-Item C:\MSTEMP1\MoveRequestStatistics.txt

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
Get-MailboxImportRequestStatistics -Identity (Get-MailboxImportRequest -Mailbox admin).RequestGuid -IncludeReport -DiagnosticInfo "showtimeslots, showtimeline, verbose" | Export-Clixml C:\MSTEMP1\ImportRequestStats.xml

$r2 = Import-Clixml C:\MSTEMP1\ImportRequestStats.xml
$r2 |fl
$r2 | select * -ExcludeProperty Report,DiagnosticInfo
$r2.Report.Failures 
$r2.Report.BadItems
$r2.Report.Connectivity
$r2.Report.Entries
$r2.Report.MailboxVerification
$r2.Report.SessionStatistics
$r2.Report.TargetMailboxSize


#endregion XML - manipulation of XML Reports


#region advanced Function
Function Test
{
[CmdletBinding(SupportsShouldProcess)] # if error inside function, it will be visible outside the function
Param(
[Parameter(Mandatory=$True, HelpMessage = "Input your name", Position=0)]
[Alias("MyName")]
[string] $name
)
$res = $PSCmdlet.ShouldContinue($name,"Title");
return $res
}

$newRes = Test -name "Victoras"
$newRes = Test -MyName "Victoras"

#endregion advanced Function

# example to show characters that are not in a set of characters using foreach
foreach ($character in [char[]]"aeioubcd") { if (@('a','e','i','o','u') -contains $character ) { continue } $character }


#endregion various examples



# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

# region log processing

get-string # to parse the log (error) 

# endregion log processing


Get-OrganizationConfig |fl *block*
help Set-OrganizationConfig -Parameter IPListBlocked

#region Calculate Hash
# Consider if to include? !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# Computes the hash value for a file by using a specified hash algorithm.
# Get-FileHash [-Path] <String[]> [-Algorithm <String> {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160} ] [ <CommonParameters>]

Get-FileHash C:\mstemp1\fstat.txt -Algorithm SHA1 | Format-List

# region - TO REIEW


$AllMbx = get-mailbox -ResultSize unlimited

$AllMbx.GetType()
$mbx= get-mailbox user9
$mbx.GetType()
$AllMbx.Contains('user9')
([System.Collections.ArrayList]$AllMbx).Contains($mbx)
$AllMbx.Contains($mbx)

$mbx=$AllMbx[0]
$mbx.GetType()
$AllMbx.Contains($mbx)


# endregion
