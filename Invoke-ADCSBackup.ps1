<#
.SYNOPSIS
    This script, when run locally on an Active Directory Certificate Services server, will perform one-time or weekday rolling backups (default) of everything necessary to restore AD CS with the exception of the CA's private key.

.DESCRIPTION
    This script REQUIRES the Microsoft ADCSAdministration PowerShell module.
    It may work with the community PSPKI module which clobbers some ADCS cmdlets but the CA template list step will fail due to replacement of Get-CATemplate.
    The Microsoft ADCSAdministration module installs with the AD CS server role by default.

    This script will back up the following to either a timestamped subfolder or a rolling weekday subfolder (default) of the specified -BackupPath:

    Critical - the script will abort if either of these fail:
    -The AD CS database and (truncated by default) logs, using Backup-CARoleService                                     #Required for CA restore, truncated logs are fine
    -The AD CS registry keys, using reg.exe export HKLM\SYSTEM\CurrentControlSet\services\CertSvc\Configuration         #CA restore is difficult without this

    Non-critical - the script will log errors and continue if any of these fail:
    -CAPolicy.inf, if it exists in the Windows directory                                                                #May not exist, likely simple to recreate
    -The CA certificate and published CRL files from the default location under Windows\System32\CertSrv\CertEnroll     #Should be available on a separate CDP/AIA server
    -The CA's Security Windows event log                                                                                #For reference only
    -An easier to read version of the AD CS registry keys, exported by certutil                                         #For reference only
    -A list of all certificate templates published on this CA only, exported by certutil                                #For reference only
    -A list of all certificate templates published in Active Directory, exported by certutil                            #For reference only
    -A full settings report of all certificate templates published in AD, exported by certutil                          #For reference only

    The AD template backups are skipped if the computer is not domain joined.
    The script logs all high-level steps and errors to a timestamped logfile in the specified -BackupPath.
    If there is a critical error that the script doesn't expect, it will attempt to leave you information in the logfile, output folder, or script directory in order.
    The script also logs its startup, completion, and all errors to a custom Windows application event log called "AD CS Backup Script".
    All items backed up include the CA's name and the date of the backup in their file or folder names. The two AD items use the domain name rather than CA name.
    Simple emails can be sent with the script's completion status using parameters documented below.

.NOTES
    Invoke-ADCSBackup.ps1
    (c) 2025 Ryan Eaton
    CC BY-SA 4.0
    https://creativecommons.org/licenses/by-sa/4.0/
    Free for commercial and non-commercial use and redistribution
    Free to modify in any way and redistribute
    Must include attribution and use the CC BY-SA 4.0 license for all redistribution and derivative work
    Author: Ryan Eaton

    VERSION HISTORY
    v1.5 - February 13, 2025
    Notification email feature has been added using the deprecated Send-MailMessage cmdlet for now
    ONLY USE WITH AN INTERNAL EMAIL SERVER.
    The email feature will be replaced in the future, probably using MailKit

    v1.1 - February 11, 2025
    PowerShell now checks for the ADCSAdministration module before allowing execution
    The script now checks that the ADCS-Cert-Authority role is installed and aborts if it is not
    Filenames of the backup output have been fine-tuned
    Email feature still not started on due to time constraints

    v1.0 - February 4, 2025
    Basic debugging and testing is complete - not all failure paths tested
    Much more optimization/best practices added
    Certutil export steps are now loops instead of repeating the same code 
    Script is fully working with no errors in a lab environment
    Ready to deploy the script to my production environment

    v0.95 - January 30, 2025
    A lot of optimization was done - testing now underway

    v0.9 - January 22, 2025
    Initial "final" version of script before testing

    TO DO
    -Parameter validation for -BackupPath
    -Separate backup validation script to take a copy of the ADCS database backup and validate it (already have code) #You can't restore a validated database file. Go figure.

    WISHLIST
    -Replace CA registry backup code with PowerShell code instead of invoking reg.exe. Looks like it might not be difficult to write code for
    -Replace the 4 non-critical backups that use certutil with CIM or .NET methods (or Win32 API calls?) if possible

    ATTRIBUTION
    Absolutely no AI tools were used at any point in the writing of this script.
    I consider the following sources to be inspirational - not foundational - to this script:

    1) A PowerShell script freely available on the website of PKI Solutions Inc. (no licensing declared):

    https://www.pkisolutions.com/blog/backing-up-adcs-certificate-authorities-part-2-of-2/

    #Script header:
    CABackup.ps1
    Scripted by: Mark B. Cooper
    PKI Solutions Inc.
    www.pkisolutions.com
    #
    Version: 1.2
    Date: October 18, 2019

    This script inspired both of my logging functions.
    Seeing it log to the Application event log, I decided to learn how to create a custom event log for my script.
    Nothing in my script is copied directly from this script.

    2) A batch file provided to me by Microsoft/JDA Partners Technical Service:

    #Script header:
    BM_DR_CABackup_v3.txt
    :: BareMetal_DisasterRecovery_CA Backup Utility with Timings - Blame: GTKT 
    :: This script is Modified and ©GTKT to build and keep copy of the CA configuration
    :: Free to use or modify as long as the Original Copyright and this notice are fully included
    :: CC BY SA 4.0

    See https://creativecommons.org/licenses/by/4.0/ for information about this license.

    Some of the certutil commands in my script are the same commands in this script, however I did not copy them directly.
    Mine are constructed and invoked using PowerShell variables and the call operator "&".
    I don't believe that a specific command to call a function of an executable can be copyrighted, since it is a function of the software e.g. "certutil.exe -ADTemplate"
    I would not consider my script to be any form of adaptation of this script. It only provided corroboration of which items are important to back up.
    This script is actually only about 15 lines of code if you remove all of the "fluff" that makes its logfile look nice.
    It has no error handling and I really wanted AD CS backup automation to be as bulletproof as I can manage to make it.

    This script also uses certutil. Per Microsoft's official documentation, certutil is not considered production-ready.
    We will avoid using it wherever we can, but it is necessary for some operations. (Maybe not if I get better at PowerShell/CIM...)
    https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/certutil

    3) Information freely available on Uwe Gradenegger's website gradenegger.eu:

    https://www.gradenegger.eu/en/create-a-backup-of-a-certification-body/ (April 2020)

    This article combined with what the previous two scripts were doing helped me decide which items to include in my script's backups.
    Uwe's name is hidden in the metadata of the documentation I received from a Microsoft PKI engagement, so I bet he was involved in script 2) as well.

    CUSTOM WINDOWS EVENT LOG EVENTS
    0-299       Information Events
    0           Event log created
    1           Script is starting
    2           Script completed with no errors

    300-399     Script Warning Events (Non-critical errors - script continues)
    300         CAPolicy.inf does not exist in $env:systemroot. It won't affect AD CS if it is not in this location
    301         Error while trying to back up CAPolicy.inf
    302         Error in certutil while trying to export one of the human-readable CA registry backups
    303         Error while trying to export one of the human-readable CA registry backups to a file
    304         Indicates that either 302 or 303 occurred at least once
    305         Error in PowerShell while trying to use certutil to export human-readable CA registry backups
    306         Error while trying to back up the default CA certificate and CRL publishing location - $env:systemroot\System32\CertSrv\CertEnroll
    307         Error while trying to back up the Windows Security event log .evtx file
    308         Error while trying to back up a list of the certificate templates published on the CA
    309         Error while retrieving the computer's domain join status or domain name. The name "GetDomainError" will be used in filenames if the AD exports still succeed
    310         Error in certutil while trying to export either a list of all certificate templates published in AD, or their full configuration details
    311         Error while trying to export either a list of all certificate templates published in AD, or their full configuration details, to a file
    312         Indicates that either 310 or 311 occurred at least once
    313         Error in PowerShell while trying to use certutil to export either a list of all certificate templates published in AD, or their full configuration details

    400-499     Email Notification Warning Events (Non-critical errors - script continues)
    400         Error while trying to create the contents of notification emails. No email will be sent.
    401         Error while trying to send a notification email.

    500-699     Script Error Events (Critical error - script aborts)
    500         Error while trying to back up the CA database and truncated logs using Backup-CARoleService
    501         Error while trying to back up the CA database and full logs using Backup-CARoleService
    502         Error in reg.exe while trying to back up the CA registry keys
    503         Error in PowerShell while trying to back up the CA registry keys
    504         Error in script logging functionality

.PARAMETER BackupPath
    Specifies the path to a folder that the script should use for all of its output.
    This is the only required parameter.
    The supplied value will either have a subfolder created for the day of the week that is then used,
    or the script will append a timestamp if the -OneTimeBackup parameter is specified.

.PARAMETER OneTimeBackup
    If specified, the script performs a one-time backup. A timestamped subfolder of the specified -BackupPath is created,
    instead of writing to rolling day-of-week subfolders.

.PARAMETER DisplayScriptLogging
    If specified, the script will write all script logging to the console.
    This is meant for interactive running of the script only.

.PARAMETER NoEventLogging
    If specified, the script does not log to the "AD CS Backup Script" custom Windows event log.
    Event logging is enabled by default.

.PARAMETER KeepFullCALogs
    If specified, the command Backup-CARoleService uses the -KeepLog parameter and does not truncate the CA logs during CA database backup.
    CA logs are truncated by default, I don't know why one might need to not truncate them. It's not necessary to restore a CA.

#The following parameters are part of the "Email" parameter set. If any one is specified then ALL must be specified.
#USE WITH INTERNAL EMAIL SERVERS ONLY.
#Your email server must be set up to allow emails to be sent this way. Exchange calls it "Anonymous Relay".
.PARAMETER SendEmailNotification
    If specified, the email notification feature is enabled.
    All 3 of the below parameters are also required.

.PARAMETER SendEmailServer
    Specifies the FQDN of the email server to send notification emails to.
    e.g. mail.woofers.org
    e.g. CARDIGAN-EX1.corgi.de

.PARAMETER SendEmailTo
    Specifies a comma-separated list of email addresses to send notification emails to.
    e.g. doggo@pets.ca, braincell@orangecat.biz, catsnake@ferrets.co.uk

.PARAMETER SendEmailFrom
    Specifies the From email address of email notifications.
    This address will receive notifications if the email is delayed or fails to send, so using a real email address is useful (but not required.)
    e.g. pkiadmin@ghostknifefish.asn.au
    e.g. monitorteam@lizard.net

.INPUTS
    None. You can't pipe objects to Invoke-ADCSBackup.ps1.

.OUTPUTS
    7 to 9 .txt files, three folders containing various files, and one each of .evtx, .log, and .reg files in a subfolder of -BackupPath.

.EXAMPLE
    PS> .\Invoke-ADCSBackup.ps1 -BackupPath C:\Temp\ADCS_Backup
    Run a default rolling day-of-week ADCS backup with default settings, event log logging, and no output to the console.

.EXAMPLE
    PS> .\Invoke-ADCSBackup.ps1 -BackupPath "E:\ADCS Backup\" -OneTimeBackup -KeepFullCALogs -DisplayScriptLogging -NoEventLogging
    Run a one-time backup of ADCS with full database logs and output to the console. Skip event log logging.

.EXAMPLE
    PS> .\Invoke-ADCSBackup.ps1 -BackupPath D:\backups -SendEmailNotification -SendEmailServer mail.clownloach.com -SendEmailTo operations@clownloach.com, archnemesis@upsidedowncatfish.or.jp -SendEmailFrom operations@clownloach.com
    Run a default rolling day-of-week ADCS backup with default settings, event log logging, and no output to the console.
    Also send an email notification of the script's completion status to operations@clownloach.com and archnemesis@upsidedowncatfish.or.jp.
    The email will be sent using the server mail.clownloach.com with the From address of operations.clownloach.com.
    Using a From address that is the same as one of the To addresses will ensure that any delayed or failed email notifications also go to that To address.
#>

[CmdletBinding(DefaultParameterSetName = "None", PositionalBinding = $false)]
#Requires -Version 5.1
#Requires -RunAsAdministrator
#Requires -Modules ADCSAdministration
Param(
    [Parameter(Mandatory, Position = 0)]
    [string]$BackupPath,
    [Parameter()]
    [switch]$OneTimeBackup,
    [Parameter()]
    [switch]$DisplayScriptLogging,
    [Parameter()]
    [switch]$NoEventLogging,
    [Parameter()]
    [switch]$KeepFullCALogs,
    [Parameter(Mandatory, ParameterSetName = "Email")]
    [switch]$SendEmailNotification,
    [Parameter(Mandatory, ParameterSetName = "Email")]
    [string]$SendEmailServer,
    [Parameter(Mandatory, ParameterSetName = "Email")]
    [string]$SendEmailTo,
    [Parameter(Mandatory, ParameterSetName = "Email")]
    [string]$SendEmailFrom
)

#FUNCTIONS
function Get-TimeStamp {
    <#
    .DESCRIPTION
        Simple timestamp function for log entries.
        Example output: [2025-02-14 15:30:44]
    #>
    return "[$(Get-Date -format 'yyyy-MM-dd HH:mm:ss' -ErrorAction Stop)]"
}

function Show-LogEntry ($LogEntry) {
    <#
    .DESCRIPTION
        Function to write various script logging to the console if the -DisplayScriptLogging parameter is specified.
        This removes a LOT of switch statements from the rest of the script.

    .PARAMETER LogEntry
        Specifies the contents of the log entry to be written to the console.
        A timestamp is added at the start of the line.
    #>
    switch ($DisplayScriptLogging) {
        $true {
            Write-Host "$(Get-TimeStamp) $LogEntry" -ErrorAction Stop
        }
        $false {
            #If -DisplayScriptLogging was not specified, do nothing
        }
    }
}

function Write-ScriptLogEntry ($LogText, [switch]$HideFromConsole = $false) {
    <#
    .DESCRIPTION
        Function to write standard logging of this script to a logfile stored in the backup destination.
        This is always on, but printing messages to the console is off by default.
        Add the -DisplayScriptLogging parameter to enable showing logs when run interactively.
        REQUIRES the Show-LogEntry function above.
        Attempts to use the Write-EventLogEntry function below if it fails,
        but if the event log doesn't exist yet, that won't work and will fail silently.

    .PARAMETER LogText
        Specifies the contents of the log entry to be written to the logfile.
        A timestamp is added at the start of the line.

    .PARAMETER HideFromConsole
        If specified, skip writing this log entry to the console even if -DisplayScriptLogging was specified for the script.
    #>
    try {
        "$(Get-TimeStamp) $LogText" | Out-File -FilePath $LogFile -Append -ErrorAction Stop
        switch ($HideFromConsole) {
            $false {
                Show-LogEntry $LogText
            }
            $true {
                #Do nothing.
            }
        }
    }
    catch {
        Show-LogEntry "Invoke-ADCSBackup.ps1 script logging failed, script cannot continue. Do you have permission to write to $($BackupPath)? Error: $($_) $($_.ScriptStackTrace) Exiting."
        try {
            Write-EventLogEntry "$(Get-TimeStamp) Invoke-ADCSBackup.ps1 script logging failed, script cannot continue. Do you have permission to write to $($BackupPath)? Error: $($_) $($_.ScriptStackTrace) Exiting." 504 "Error"
        }
        catch {
            #Do nothing. Depending on where we are in the script, this may not work so we are just silencing any error output
        }
        finally {
            throw "$(Get-TimeStamp) Invoke-ADCSBackup.ps1 script logging failed, script cannot continue. Do you have permission to write to $($BackupPath)? Error: $($_) $($_.ScriptStackTrace) Exiting."
            #Thrown either to the catch of THE BIG TRY CATCH, or to the nested catch of the finally block of THE BIG TRY CATCH
        }
    }
}

function Write-EventLogEntry ($EventText, $EventID, $EventType) {
    <#
    .DESCRIPTION
        Function to write to a custom Windows event log for this script (enabled by default.)
        This can be disabled via the -NoEventLogging parameter.
        The parameter variables are typed for this function,
        just in case the Write-EventLog cmdlet might not play nice with other input types.
        REQUIRES the Show-LogEntry function above.
        Attempts to use the Write-ScriptLogEntry function above if it fails,
        which will fail silently if it doesn't work.

    .PARAMETER EventText
        Specifies the main contents of the event to be written to the event log.

    .PARAMETER EventID
        Specifies the event ID of the log entry to be written to the event log.

    .PARAMETER EventType
        Specifies the type of event log entry to be written to the event log.
        Examples: "Information", "Warning", "Error"
    #>
    switch ($NoEventLogging) {
        $false {
            try {
                Write-EventLog -LogName "AD CS Backup Script" -Source "ADCSBackup" -EventId $EventID -EntryType $EventType -Message $EventText -Category 0 -ErrorAction Stop
            }
            catch {
                Show-LogEntry "Invoke-ADCSBackup.ps1 event logging failed, script cannot continue. Do you have permission to write to $($env:systemroot)\System32\Winevt\Logs? Error: $($_) $($_.ScriptStackTrace) Exiting."
                try {
                    Write-ScriptLogEntry "$(Get-TimeStamp) Invoke-ADCSBackup.ps1 event logging failed, script cannot continue. Do you have permission to write to $($env:systemroot)\System32\Winevt\Logs? Error: $($_) $($_.ScriptStackTrace) Exiting."
                }
                catch {
                    #Do nothing. Depending on where we are in the script, this may not work. We are just silencing any error output.
                }
                finally {
                    throw "$(Get-TimeStamp) Invoke-ADCSBackup.ps1 event logging failed, script cannot continue. Do you have permission to write to $($env:systemroot)\System32\Winevt\Logs? Error: $($_) $($_.ScriptStackTrace) Exiting."
                    #Thrown either to the catch of THE BIG TRY CATCH, or to the nested catch of the finally block of THE BIG TRY CATCH
                }
            }
        }
        $true {
            #If -NoEventLogging specified, do nothing
        }
    }
}

function Add-ArrayListObject ($ArrayList, $ObjectToAdd) {
    <#
    .DESCRIPTION
        Function to add an object to an ArrayList,
        because we do this over and over again until the logfile is created and then to store any non-critical errors.

    .PARAMETER ArrayList
        Specifies which ArrayList variable to add the object to.

    .PARAMETER ObjectToAdd
        Specifies the object to be added to the specified ArrayList variable.
    #>
    try {
        [void]$ArrayList.Add($ObjectToAdd)
    }
    catch {
        Show-LogEntry "Error adding a pre-logfile log entry or non-critical script error to a list. Script logging may be incomplete. Continuing..."
    }
}

function Set-EmailNotificationContents ($ScriptStatus) {
    <#
    .DESCRIPTION
        Function to build the notification emails that the script can send.
        They need to be built inside of a function so that we can call the function once the required variables have been set.
        All variables that don't already exist are set using the script: variable scope so that they are available outside of the function.
    #>
    switch ($SendEmailNotification) {
        $false {
            #Do nothing, no need to assemble the emails
        }
        $true {
            try {
                #Set timestamps for the email
                $script:EmailStartTimeStamp = "$($ScriptInvocationTime.Split("_")[0]) $($($ScriptInvocationTime).Split("_")[1].Replace("-",":"))"   #Take a copy of the script's already-set start time, converted back to a format you would use in an email
                $script:EmailEndTimeStamp = Get-Date -format 'yyyy-MM-dd HH:mm:ss' -ErrorAction Stop                                                #When this function is called, the script is in the process of finalizing logs and exiting so this time is close enough

                #Set the contents of the 3 possible emails this script can send
                switch ($ScriptStatus) {
                    #The indentation formatting below is cursed, but it ensures that all indentations are correct in the final $EmailNotificationBody below
                    #It's HTML so the indentation doesn't actually matter to what you'll see in your email client, but...nyeh
                    #Note that's no indent on the first line and two on any subsequent lines
                    #Note the here-string @"<string>"@ below. The closing "@ CANNOT have any preceding spaces or the here-string will break
                    #It's ugly, but it works
                    "CriticalError" {
                        $script:EmailNotificationSubjectDynamic = "CRITICAL ERROR: AD CS Backup Script - FAILED"
                        $script:EmailNotificationPriorityDynamic = "High"
                        $script:EmailNotificationDynamic = @"
<p>A <span style="font-weight: bold; color: red;">CRITICAL ERROR</span> has occurred in the AD CS Backup Script on $ComputerName.</p>
        <p style="font-weight: bold; color: red;">The backup failed to complete!</p>
        <p style="font-weight: bold;">The script logfile will be attached to this email unless the error prevented its creation.</p>
        <p style="font-weight: bold; color: red;">Please take immediate action to identify and remediate this issue.</p>
"@
                    }
                    "NonCriticalError" {
                        $script:EmailNotificationSubjectDynamic = "ERROR: AD CS Backup Script - Non-Critical"
                        $script:EmailNotificationPriorityDynamic = "High"
                        $script:EmailNotificationDynamic = @"
<p>A <span style="font-weight: bold; color: orange;">non-critical error</span> has occurred in the AD CS Backup Script on $ComputerName.</p>
        <p style="font-weight: bold;">The script logfile should be attached to this email. If not, it should be in the backup folder.</p>
        <p style="font-weight: bold;">Please take action to identify and remediate this issue.</p>
"@
                    }
                    "Success" {
                        $script:EmailNotificationSubjectDynamic = "SUCCESS: AD CS Backup Script - Complete"
                        $script:EmailNotificationPriorityDynamic = "Normal"
                        $script:EmailNotificationDynamic = @"
<p style="font-weight: bold; color:green;">The AD CS Backup Script has completed successfully on $ComputerName.</p>
        <p>The script logfile should be attached to this email. If not, it should be in the backup folder.</p>
"@
                    }
                }

                #Now we assemble the full final email
                #Again, weird here-string formatting as above
                $script:EmailNotificationSubject = "$($EmailNotificationSubjectDynamic) [$($EmailStartTimeStamp)]"
                $script:EmailNotificationPriority = $EmailNotificationPriorityDynamic
                $script:EmailNotificationBody = @"
<html>

<body>
    <div style="font-family: 'Aptos', sans-serif; font-size: 12.0pt;">
        <p>Hello,</p>
        $($EmailNotificationDynamic)
        <p>Start time: $($EmailStartTimeStamp)<br />End time: $($EmailEndTimeStamp)</p>
        <p>Thanks,</p>
        <p>AD CS Backup Script</p>
    </div>
</body>

</html>
"@

            }
            catch {
                Write-ScriptLogEntry "Error: -SendEmailNotification was specified, but something went wrong creating the emails. No email will be sent. Continuing..."
                Write-EventLogEntry "Error: -SendEmailNotification was specified, but something went wrong creating the emails. No email will be sent. Continuing..." 400 "Warning"
                #Switch the email feature off so the script doesn't try to send emails that were not built
                $script:SendEmailNotification = $false
            }
        }
    }
}

function Send-EmailNotification ($Subject, $Body, $Priority) {
    <#
    .DESCRIPTION
        Function to send simple notification emails via an internal email server using "Anonymous Relay" as Exchange calls it.
        DO NOT USE WITH EXTERNAL MAIL SERVERS. This is intended for on-premise Exchange. Other email server software might work.
        Send-MailMessage has been deprecated by Microsoft because it can't use modern security methods.
        Note that -UseSsl is specified in this function. Your email server must support SSL.
        Note that -BodyAsHtml is specified in this function. This script contains pre-built HTML notification emails below.

    .PARAMETER Subject
        Specifies the contents of the subject line of the email.

    .PARAMETER Body
        Specifies the contents of the body of the email.
        This is best passed within a here-string, which can handle multiple lines and HTML: @"<string>"@

    .PARAMETER Priority
        Specifies what "Priority" value to mark the email with.
    #>
    switch ($SendEmailNotification) {
        $false {
            #Do nothing
        }
        $true {
            try {
                Send-MailMessage -SmtpServer $SendEmailServer -To $SendEmailTo -From $SendEmailFrom -Attachments $LogFile -Subject $Subject -Body $Body -DeliveryNotificationOption Delay,OnFailure -Priority $Priority -BodyAsHtml -UseSsl -ErrorAction Stop
            }
            catch {
                Add-ArrayListObject $NonCriticalErrors @{"Error" = "$(Get-TimeStamp) Script has finished. -SendEmailNotification was specified, but sending email failed! Exiting..." ; "Location" = $_.ScriptStackTrace}
                Show-LogEntry "Script has finished. -SendEmailNotification was specified, but sending email failed! Exiting..."
                Write-EventLogEntry "Script has finished. -SendEmailNotification was specified, but sending email failed! Continuing..." 401 "Warning"
            }
        }
    }
}

#THE BIG TRY CATCH
#We need to catch any critical errors to handle at the #WRAP UP/finally block at the end of the script.
try {

    #QUICK SANITY CHECK
    #Does this computer have the ADCS server role installed?
    if ($null -eq (Get-WindowsFeature -Name "ADCS-Cert-Authority" -ErrorAction Stop).InstallState) {
        Show-LogEntry "Active Directory Certificate Services is not installed on this PC. Script cannot continue. Exiting."
        throw "$(Get-TimeStamp) Active Directory Certificate Services is not installed on this PC. Script cannot continue. Exiting."
        #Thrown to the end of THE BIG TRY CATCH
    }

    #SETUP
    #Initialize empty variables/arrays to be populated with error or log objects, so we can include that information in an email notification/insert log entries from before the logfile has been created
    #Specified with "script:" variable scope to ensure accessibility by all child scopes such as try catch statements
    $script:CriticalErrorOccurred = $null                                       #We don't need an array because the script exits as soon as any critical error occurs
    $script:NonCriticalErrors = [System.Collections.ArrayList]::new()           #We use this to capture all non-critical errors for output later
    $script:PreLogFileLogEntries = [System.Collections.ArrayList]::new()        #We use this to capture log entries from before the logfile exists and then insert them into the logfile

    #Initialize variables used later in the script and set them to the "script" variable scope
    try {
        $script:ScriptVersion = "1.5"                                                                                       #Set script version for identification in logs
        $script:ScriptInvocationTime = Get-Date -format 'yyyy-MM-dd_HH-mm-ss' -ErrorAction Stop                             #Getting the start time once ensures that the folder and logfile name are the same
        $script:ScriptInvocationDate = Get-Date -format 'yyyy-MM-dd' -ErrorAction Stop                                      #We need this several times, no use processing it repeatedly
        $script:ScriptInvocationShortDayOfWeek = $((Get-Date -ErrorAction Stop).DayOfWeek) -replace '^(.{0,3}).*', '$1'     #We only use this once, but I think it's easier to parse here
        $script:CustomEventLogDefaultPath = "$($env:systemroot)\System32\Winevt\Logs\AD CS Backup Script.evtx"              #Changing this will break the script unless you make heavy modifications
        $script:CaPolicyInfPath = "$($env:systemroot)\capolicy.inf"                                                         #You can change this, but capolicy.inf only affects AD CS if it is in this location
        $script:RegPath = "$($env:systemroot)\System32\reg.exe"                                                             #Path to reg.exe
        $script:CertutilPath = "$($env:systemroot)\System32\certutil.exe"                                                   #Path to certutil.exe
        $script:ComputerName = "$([System.Environment]::MachineName)"                                                       #Used in email notifications
        $script:CAName = $null                                                                                              #We will set this in the next code block
        $script:WorkingBackupPath = $null                                                                                   #We will set this two code blocks down
        $script:LogFile = $null                                                                                             #We will set this two code blocks down
        try {
            if ((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).PartOfDomain) {                         #Returns $true if domain joined, $false if not
                $script:CADomain = "$((Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Domain)"          #Get domain name for use in filenames
            }
            else {
                $script:CADomain = $null                                                                                    #We'll use this $null later to skip AD backups
            }
        }
        catch {
            $script:CADomain = "GetDomainError"                                                                             #We'll use this to still attempt AD template backup
        }
    }
    catch {
        Show-LogEntry "Error initializing variables. Is the Get-Date cmdlet working? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        throw "$(Get-TimeStamp) Error initializing variables. Is the Get-Date cmdlet working? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        #Thrown to the end of THE BIG TRY CATCH
    }

    #Grab the CA's name and replace spaces with dashes for easy use in filenames. If this goes wrong, make it obvious to see later and at least attempt to perform the rest of the backup
    try {
        Show-LogEntry "Retrieving CA name from the registry for use in filenames..."
        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Retrieving CA name from the registry for use in filenames..."
        $CAName = ((Get-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\Services\CertSvc\Configuration -ErrorAction Stop).Active).Replace(" ", "-")
        Show-LogEntry "Successfully retrieved `"$(($CAName).Replace("-"," "))`" from the registry."
        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Successfully retrieved `"$(($CAName).Replace("-"," "))`" from the registry."
    }
    catch {
        Show-LogEntry "Retrieving CA name from the registry failed. Using `"CANameRegQueryFailed`" as CA name for filenames and continuing..."
        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Retrieving CA name from the registry failed. Using `"CANameRegQueryFailed`" as CA name for filenames and continuing..."
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        $CAName = "CANameRegQueryFailed"
    }

    #Set "real" backup folder path as a subfolder of $BackupPath and set logfile name and path based on specified parameters
    #Transferring $BackupPath to another variable also prevents $BackupPath from being modified and re-validated by PowerShell which can break the script (validation not implemented yet)
    switch ($OneTimeBackup) {
        $true {
            $script:WorkingBackupPath = "$($BackupPath.TrimEnd("\"))\$($CAName)_Backup-$($ScriptInvocationTime)"
            $LogFile = "$($WorkingBackupPath)\$($CAName)_Backup-$($ScriptInvocationTime).log"
        }
        $false {
            $script:WorkingBackupPath = "$($BackupPath.TrimEnd("\"))\$($ScriptInvocationShortDayOfWeek)"
            $LogFile = "$($WorkingBackupPath)\$($CAName)_Backup-$($ScriptInvocationTime).log"
        }
    }
        
    #Create the backup folder if it does not exist
    try {
        if (-not (Test-Path -Path $WorkingBackupPath -ErrorAction Stop)) {
            Show-LogEntry "$($WorkingBackupPath) does not exist, creating it..."
            Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) $($WorkingBackupPath) does not exist, creating it..."
            New-Item $WorkingBackupPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Show-LogEntry "$($WorkingBackupPath) created successfully."
            Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) $($WorkingBackupPath) created successfully."
        }
    }
    catch {
        Show-LogEntry "Error creating $($WorkingBackupPath). Do you have permissions to create a folder here? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        throw "$(Get-TimeStamp) Backup folder $($WorkingBackupPath) creation failed. Do you have write permission to its parent folder? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        #Thrown to the end of THE BIG TRY CATCH
    }

    #If rolling backups are enabled and this weekday's subfolder exists, wipe its contents (rolling backups are enabled by default)
    switch ($OneTimeBackup) {
        $false {
            try {
                #Using * in the path will return $true if any (non-hidden and non-system) files or folders exist in the specified parent folder, and $false if not
                if (Test-Path "$($WorkingBackupPath)\*" -ErrorAction Stop) {
                    Show-LogEntry "$($WorkingBackupPath) contains data, likely from a previous week's backup. Deleting all contents of $($WorkingBackupPath)..."
                    Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) $($WorkingBackupPath) contains data, likely from a previous week's backup. Deleting all contents of $($WorkingBackupPath)..."
                    try {
                        Remove-Item "$($WorkingBackupPath)\*" -Recurse -Force -ErrorAction Stop
                        Show-LogEntry "Previous $((Get-Date).DayOfWeek)'s backup folder contents deleted successfully."
                        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Previous $((Get-Date).DayOfWeek)'s backup folder contents deleted successfully."
                    }
                    catch {
                        #Append the date to the weekday folder name, create that folder, and try to use it instead
                        $WorkingBackupPath = "$($WorkingBackupPath)_$($ScriptInvocationDate)"
                        #Modify the logfile path with the new folder name
                        $LogFile = "$($WorkingBackupPath)\$($CAName)_Backup-$($ScriptInvocationTime).log"
                        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
                        Show-LogEntry "Error removing previous $((Get-Date).DayOfWeek)'s backup folder contents. Attempting to recover by creating an alternate folder..."
                        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Error removing previous $((Get-Date).DayOfWeek)'s backup folder contents. Attempting to recover by creating an alternate folder..."
                        try {
                            if (Test-Path -Path $WorkingBackupPath -ErrorAction Stop) {
                                $RecoveryFolderAlreadyExists = "A recovery attempt folder for $($ScriptInvocationDate) already exists. Attempting to delete its contents. This may not work..."
                                Add-ArrayListObject $NonCriticalErrors @{"Error" = $RecoveryFolderAlreadyExists ; "Location" = "$($MyInvocation.ScriptName): line $($MyInvocation.ScriptLineNumber)"}
                                Show-LogEntry $RecoveryFolderAlreadyExists
                                Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) $($RecoveryFolderAlreadyExists)"
                                Remove-Item "$($WorkingBackupPath)\*" -Recurse -Force -ErrorAction Stop
                            }
                            else {
                                New-Item $WorkingBackupPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            }
                        }
                        catch {
                            Show-LogEntry "Error attempting to recover by creating an alternate folder. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                            throw "$(Get-TimeStamp) Error attempting to recover by creating an alternate folder. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                            #Thrown to the catch block below
                        }
                    }
                }
            }
            catch {
                Show-LogEntry "Error testing for a previous $((Get-Date).DayOfWeek)'s backup folder. Do you have read permissions to $($WorkingBackupPath)? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                throw "$(Get-TimeStamp) Error testing for a previous $((Get-Date).DayOfWeek)'s backup folder. Do you have read permissions to $($WorkingBackupPath)? Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                #Thrown to the end of THE BIG TRY CATCH
            }
        }
        $true {
            Show-LogEntry "-OneTimeBackup parameter was specified. Backing up to $($WorkingBackupPath)."
            Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) -OneTimeBackup parameter was specified. Backing up to $($WorkingBackupPath)."
        }
    }

    #Create the logfile (this should always have a unique name as the timestamp is to the second)
    try {
        Show-LogEntry "Creating $($LogFile)..."
        Add-ArrayListObject $PreLogFileLogEntries "$(Get-TimeStamp) Creating $($LogFile)..."
        New-Item $LogFile -ItemType File -Force -ErrorAction Stop | Out-Null
        Write-ScriptLogEntry "$($LogFile) has been created successfully."
        try {
            Write-ScriptLogEntry "Below is all output from the script before this logfile's creation:" -HideFromConsole
            $PreLogFileLogEntries | ForEach-Object {[PSCustomObject]$_} | Out-File -FilePath $LogFile -Append -Force -ErrorAction Stop
            Write-ScriptLogEntry "End of pre-logfile log entries. Continuing..." -HideFromConsole
        }
        catch {
            Write-ScriptLogEntry "Error in collection of pre-logfile log entries or in writing them to $($LogFile). Continuing anyway..."
        }
    }
    catch {
        Show-LogEntry "Logfile $($LogFile) creation failed. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        throw "$(Get-TimeStamp) Logfile $($LogFile) creation failed. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
        #Thrown to the end of THE BIG TRY CATCH
    }

    #Register a new event log source if it does not exist, if event logging is enabled (enabled by default)
    switch ($NoEventLogging) {
        $false {
            try {
                if (-not (Test-Path -Path $CustomEventLogDefaultPath)) {
                    Write-ScriptLogEntry "Event log `"AD CS Backup Script`" does not exist. Creating it..."
                    New-EventLog -LogName "AD CS Backup Script" -Source "ADCSBackup" -ErrorAction Stop
                    Limit-EventLog -LogName "AD CS Backup Script" -MaximumSize 51200KB -ErrorAction SilentlyContinue
                    Write-ScriptLogEntry "Event log `"AD CS Backup Script`" has been created and source `"ADCSBackup`" registered."
                    Write-EventLogEntry "Event log `"AD CS Backup Script`" has been created and source `"ADCSBackup`" registered." 0 "Information"
                }
            }
            catch {
                Write-ScriptLogEntry "Event log creation failed. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                throw "$(Get-TimeStamp) Event log creation failed. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                #Thrown to the end of THE BIG TRY CATCH
            }
        }
        $true {
            Write-ScriptLogEntry "-NoEventLogging parameter was specified. Logging to a custom Windows Event Log is disabled. Continuing..."
        }
    }

    #MAIN SCRIPT
    Write-ScriptLogEntry "Script starting... AD CS backup script version $($ScriptVersion). Writing logs to $($LogFile)."
    switch ($NoEventLogging) {
        $false {
            Write-ScriptLogEntry "Also writing logs to the custom `"AD CS Backup Script`" Windows event log. It will be created if it does not exist."
        }
        $true {
            #Do nothing.
        }
    }
    Write-ScriptLogEntry "Script is running as $([System.Environment]::UserDomainName)\$([System.Environment]::UserName)."
    Write-EventLogEntry "Script starting... AD CS backup script version $($ScriptVersion). Writing logs to $($LogFile). Running as $([System.Environment]::UserDomainName)\$([System.Environment]::UserName). Please note that this event log is not as detailed as $($LogFile)." 1 "Information"

    #The database/logs backup and the registry backup are critical. The script will abort if either fail
    #Back up the CA database and logs only, and NOT the CA's private key. NEVER back up the CA's private key with this type of backup. CA logs are truncated by default
    switch ($KeepFullCALogs) {
        $false {
            Write-ScriptLogEntry "Backing up the CA database and truncated CA logs..."
            Write-ScriptLogEntry "Backup-CARoleService output below:" -HideFromConsole
            try {
                New-Item -Path "$($WorkingBackupPath)" -Name "$($CAName)_DataBase_$($ScriptInvocationDate)" -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Backup-CARoleService -Path "$($WorkingBackupPath)\$($CAName)_DataBase_$($ScriptInvocationDate)" -DatabaseOnly -Verbose -ErrorAction Stop *>> $LogFile
                Write-ScriptLogEntry "CA database and truncated logs backup completed."
            }
            catch {
                Write-ScriptLogEntry "Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                Write-EventLogEntry "Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)" 500 "Error"
                throw "$(Get-TimeStamp) Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
            }
        }
        $true {
            Write-ScriptLogEntry "The -KeepFullCALogs parameter is set. Backing up the CA database and full CA logs..."
            Write-ScriptLogEntry "Backup-CARoleService output below:" -HideFromConsole
            try {
                New-Item -Path "$($WorkingBackupPath)" -Name "$($CAName)_DataBase-FullLogs_$($ScriptInvocationDate)" -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Backup-CARoleService -Path "$($WorkingBackupPath)\$($CAName)_DataBase-FullLogs_$($ScriptInvocationDate)" -DatabaseOnly -KeepLog -Verbose -ErrorAction Stop *>> $LogFile
                Write-ScriptLogEntry "CA database and full logs backup completed."
            }
            catch {
                Write-ScriptLogEntry "Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
                Write-EventLogEntry "Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)" 501 "Error"
                throw "$(Get-TimeStamp) Error backing up CA database and logs. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
            }
        }
    }

    #Back up the AD CS registry keys from the CA
    Write-ScriptLogEntry "Backing up the AD CS registry keys from the CA..."
    try {
        $RegParameters = @("export", "HKLM\SYSTEM\CurrentControlSet\services\CertSvc\Configuration", "`"$($WorkingBackupPath)\$($CAName)_CertSvc-Registry-Config_$($ScriptInvocationDate).reg`"")
        & $RegPath $RegParameters *>> $null
        if ($LASTEXITCODE -ne 0) {
            #This throw will jump execution to the catch statement below while passing the exit code from reg.exe
            throw $LASTEXITCODE
        }
        Write-ScriptLogEntry "AD CS registry key backup completed."
    }
    catch {
        #If reg.exe fails, the error object thrown above should match the regex below with an exit code between 1 and 255. So pass the reg.exe error as the script aborts
        #Else, this should mean the error was in PowerShell and not reg.exe so pass whatever error PowerShell sees as the script aborts
        if ($_.Exception -match '^System.Management.Automation.RuntimeException: [0-9]{1,3}$') {
            Write-ScriptLogEntry "Error in reg.exe backing up CA registry. Script cannot continue. Exiting. Error: $($LASTEXITCODE)"
            Write-EventLogEntry "Error in reg.exe backing up CA registry. Script cannot continue. Exiting. Error: $($LASTEXITCODE)" 502 "Error"
            throw "$(Get-TimeStamp) Error in reg.exe backing up CA registry. Script cannot continue. Exiting. Error: $($LASTEXITCODE)"
            #Thrown to the end of THE BIG TRY CATCH
        }
        else {
            Write-ScriptLogEntry "Error in PowerShell backing up CA registry. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
            Write-EventLogEntry "Error in PowerShell backing up CA registry. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)" 503 "Error"
            throw "$(Get-TimeStamp) Error in PowerShell backing up CA registry. Script cannot continue. Exiting. Error: $($_) $($_.ScriptStackTrace)"
            #Thrown to the end of THE BIG TRY CATCH
        }
    }

    #All backups past this point are considered non-critical and will not abort the script if they fail
    #Back up the CAPolicy.inf file if it exists
    Write-ScriptLogEntry "Backing up CAPolicy.inf..." 
    try {
        if (Test-Path -Path $CaPolicyInfPath -ErrorAction Stop) {
            #Copy CAPolicy.inf to a named subfolder, but if that fails then just try to copy it to $WorkingBackupPath
            try {
                New-Item -Path $WorkingBackupPath -Name "$($CAName)_CAPolicy-inf_$($ScriptInvocationDate)" -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Copy-Item -Path $CaPolicyInfPath -Destination "$($WorkingBackupPath)\$($CAName)_CAPolicy-inf_$($ScriptInvocationDate)" -Force -ErrorAction Stop
            }
            catch {
                Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
                Write-ScriptLogEntry "Error backing up CAPolicy.inf to a named subfolder, attempting to back up directly to $($WorkingBackupPath)..."
                Copy-Item -Path $CaPolicyInfPath -Destination $WorkingBackupPath -Force -ErrorAction Stop
            }
            Write-ScriptLogEntry "CAPolicy.inf backup completed."
        }
        else {
            Write-ScriptLogEntry "CAPolicy.inf not found in $($env:systemroot). It must exist in this location in order to be used by AD CS. Continuing..."
            Write-EventLogEntry "CAPolicy.inf not found in $($env:systemroot). It must exist in this location in order to be used by AD CS. Continuing..." 300 "Warning"
        }
    }
    catch {
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        Write-ScriptLogEntry "Error backing up CAPolicy.inf. Continuing... Error: $($_) $($_.ScriptStackTrace)"
        Write-EventLogEntry "Error backing up CAPolicy.inf. Continuing... Error: $($_) $($_.ScriptStackTrace)" 301 "Warning"
    }

    #Back up the CA certificate and CRL files from the default publishing location
    Write-ScriptLogEntry "Backing up the CA certificate and CRL files..."
    try {
        if ((Test-Path "$($env:systemroot)\System32\CertSrv\CertEnroll\*" -Include *.crt -ErrorAction Stop) -or (Test-Path "$($env:systemroot)\System32\CertSrv\CertEnroll\*" -Include *.crl -ErrorAction Stop)) {
            #Copy the files to a named subfolder, but if that fails then just try to copy them to $WorkingBackupPath
            try {
                New-Item -ItemType Directory -Path $WorkingBackupPath -Name "$($CAName)_CertEnroll_$($ScriptInvocationDate)" -Force -ErrorAction Stop | Out-Null
                Copy-Item -Path "$($env:systemroot)\System32\CertSrv\CertEnroll" -Destination "$($WorkingBackupPath)\$($CAName)_CertEnroll_$($ScriptInvocationDate)" -Recurse -Force -ErrorAction Stop | Out-Null
            }
            catch {
                Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
                Write-ScriptLogEntry "Error backing up the CertEnroll folder to a named subfolder. Attempting to back up CA certificate and CRL files directly to $($WorkingBackupPath)..."
                Copy-Item -Path "$($env:systemroot)\System32\CertSrv\CertEnroll" -Destination $WorkingBackupPath -Recurse -Force -ErrorAction Stop | Out-Null
            }
            Write-ScriptLogEntry "CA certificate and CRL files backup completed."
        }
        else {
            Write-ScriptLogEntry "The $($env:systemroot)\System32\CertSrv\CertEnroll folder contains no .crt or .crl files. No files have been backed up. Continuing..."
            throw "$(Get-TimeStamp) The $($env:systemroot)\System32\CertSrv\CertEnroll folder contains no .crt or .crl files. No files have been backed up. Continuing..."
            #Thrown to the catch statement below
        }
    }
    catch {
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        Write-ScriptLogEntry "Error backing up the CertEnroll folder containing the CA certificate and CRL files. Do you have permission to? Continuing... Error: $($_) $($_.ScriptStackTrace)"
        Write-EventLogEntry "Error backing up the CertEnroll folder containing the CA certificate and CRL files. Do you have permission to? Continuing... Error: $($_) $($_.ScriptStackTrace)" 306 "Warning"
    }
    
    #Back up the CA Security event logs
    Write-ScriptLogEntry "Backing up the CA Security event log..."
    try {
        $SecurityEventLog = Get-CimInstance Win32_NTEventlogFile -ErrorAction Stop | Where-Object LogfileName -eq "Security" -ErrorAction Stop
        Invoke-CimMethod -InputObject $SecurityEventLog -MethodName BackupEventlog -Arguments @{ArchiveFileName = "$($WorkingBackupPath)\$($CAName)_Security-Event-Log_$($ScriptInvocationDate).evtx"} -ErrorAction Stop | Out-Null
        Write-ScriptLogEntry "CA Security event log backup completed."
    }
    catch {
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        Write-ScriptLogEntry "Error backing up the CA Security event log. Continuing... Error: $($_) $($_.ScriptStackTrace)"
        Write-EventLogEntry "Error backing up the CA Security event log. Continuing... Error: $($_) $($_.ScriptStackTrace)" 307 "Warning"
    }

    #Back up a more human-readable version of the AD CS registry configuration from the CA
    Write-ScriptLogEntry "Backing up AD CS registry key config reports from the CA..."
    try {
        #Create an array of hashtables so we can loop this
        $CertutilRegParams = @(
            @{"Params" = @("-v", "-getreg") ; "Component" = "CertSvc"}                              #To get basic CertSvc configuration
            @{"Params" = @("-v", "-getreg", "CA") ; "Component" = "CA"}                             #To get basic CA configuration
            @{"Params" = @("-v", "-getreg", "CA\CSP") ; "Component" = "CA-Certificate"}             #To get CA certificate configuration
            @{"Params" = @("-v", "-getreg", "CA\EncryptionCSP") ; "Component" = "CA-Encryption"}    #To get CA encryption configuration
            @{"Params" = @("-v", "-getreg", "CA\PolicyModules") ; "Component" = "Policy-Modules"}   #To get policy module configuration
            @{"Params" = @("-v", "-getreg", "CA\ExitModules") ; "Component" = "Exit-Modules"}       #To get exit modules configuration
        )
        $CertUtilRegOutFiles = @()                                                                  #To hold a list of output filenames to check for later
        foreach ($Command in $CertutilRegParams) {
            #We have to use certutil for this
            $RegOutFile = "$($WorkingBackupPath)\$($CAName)_$($Command.Component)-Config_$($ScriptInvocationDate).txt"
            $CertUtilRegOutFiles += $RegOutFile
            [System.Collections.ArrayList]$CertUtilRegOutput = & $CertutilPath $($Command.Params)
            if ($LASTEXITCODE -ne 0) { 
                Add-ArrayListObject $NonCriticalErrors  @{"Error" = "$($LASTEXITCODE)" ; "Location" = "$($MyInvocation.ScriptName): line $($MyInvocation.ScriptLineNumber)"}
                $CertUtilRegOutput = $null
                $RegOutFile = $null
                Write-ScriptLogEntry "Error in certutil backing up $($Command.Component) configuration. Continuing... Error: $($LASTEXITCODE)"
                Write-EventLogEntry "Error in certutil backing up $($Command.Component) configuration. Continuing... Error: $($LASTEXITCODE)" 302 "Warning"
                continue
            }
            #Remove the final line that indicates certutil completed successfully, then output to a file
            $CertUtilRegOutput.RemoveAt($($CertUtilRegOutput.Count) - 1)
            try {
                $CertUtilRegOutput > $RegOutFile
            }
            catch {
                Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
                $CertUtilRegOutput = $null
                $RegOutFile = $null
                Write-ScriptLogEntry "Error writing $($Command.Component) configuration to file. Continuing... Error: $($_) $($_.ScriptStackTrace)"
                Write-EventLogEntry "Error writing $($Command.Component) configuration to file. Continuing... Error: $($_) $($_.ScriptStackTrace)" 303 "Warning"
                continue
            }
            $CertUtilRegOutput = $null
            $RegOutFile = $null
            Write-ScriptLogEntry "Backup of $($Command.Component) configuration completed."
        }
        if (($CertUtilRegOutFiles | Test-Path) -notcontains $false) {
            #If all of the output files exist, there were no errors
            Write-ScriptLogEntry "Backup of easy-to-read versions of the AD CS registry keys from the CA completed."
        }
        else {
            #Else, at least one of the files was not created
            Write-ScriptLogEntry "Backup of easy-to-read versions of the AD CS registry keys from the CA completed with errors. The errors should be shown above this message. Continuing..."
            Write-SEventLogEntry "Backup of easy-to-read versions of the AD CS registry keys from the CA completed with errors. The errors should be shown above this message. Continuing..." 304 "Warning"
        }
    }
    catch {
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        Write-ScriptLogEntry "Error in PowerShell backing up easy-to-read AD CS registry information. Continuing... Error: $($_) $($_.ScriptStackTrace)"
        Write-EventLogEntry "Error in PowerShell backing up easy-to-read AD CS registry information. Continuing... Error: $($_) $($_.ScriptStackTrace)" 305 "Warning"
    }

    #Back up a list of the certificate templates that are published on the CA
    Write-ScriptLogEntry "Backing up a list of the certificate templates that are published on this CA..."
    try {
        #If you happen to have installed the open-source PSPKI PowerShell module, it overwrites the Get-CATemplate command with one that will not work with these parameters
        Get-CATemplate -ErrorAction Stop | Format-List -ErrorAction Stop | Out-File -FilePath "$($WorkingBackupPath)\$($CAName)_PublishedTemplatesList_$($ScriptInvocationDate).txt" -Encoding String -Force -ErrorAction Stop
        Write-ScriptLogEntry "Backing up a list of the certificate templates published on this CA completed."
    }
    catch {
        Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
        Write-ScriptLogEntry "Error backing up a list of the certificate templates published on this CA. If you have installed the community PSPKI PowerShell module, this will not work. Continuing... Error: $($_) $($_.ScriptStackTrace)"
        Write-EventLogEntry "Error backing up a list of the certificate templates published on this CA. If you have installed the community PSPKI PowerShell module, this will not work. Continuing... Error: $($_) $($_.ScriptStackTrace)" 308 "Warning"
    }

    if ($null -ne $CADomain) {
        #Back up a list of and the full configuration details of all the certificate templates that are published in AD
        Write-ScriptLogEntry "Backing up a list of and the full configuration details of all certificate templates published in AD..."
        if ($CADomain -eq "GetDomainError") {
            Write-ScriptLogEntry "There was an error retrieving this computer's domain join status or domain name. The name `"GetDomainError`" will be used if the two following AD exports succeed..."
            Write-EventLogEntry "There was an error retrieving this computer's domain join status or domain name. The name `"GetDomainError`" will be used if the two following AD exports succeed..." 309 "Warning"
        }
        try {
            #We have to use certutil for this
            #Create an array of hashtables so we can loop this
            $CertutilADTemplateParams = @(
                @{"Params" = @("-ADTemplate") ; "Descriptions" = @("ADTemplates-List", "a list")}                                               #To get a list of all templates published in AD
                @{"Params" = @("-v", "-ADTemplate") ; "Descriptions" = @("ADTemplates-FullConfiguration", "the full configuration details")}    #To get the full configuration details of all templates published in AD
            )
            $CertutilADTemplateOutFiles = @()                                                                                                   #To hold a list of output filenames to check for later
            foreach ($Command in $CertutilADTemplateParams) {
                #We have to use certutil for this
                $ADTemplateOutFile = "$($WorkingBackupPath)\$($CADomain)_$($Command.Descriptions[0])_$($ScriptInvocationDate).txt"
                $CertutilADTemplateOutFiles += $ADTemplateOutFile
                [System.Collections.ArrayList]$CertUtilADTemplateOutput = & $CertutilPath $($Command.Params)
                if ($LASTEXITCODE -ne 0) { 
                    Add-ArrayListObject $NonCriticalErrors @{"Error" = "$($LASTEXITCODE)" ; "Location" = "$($MyInvocation.ScriptName): line $($MyInvocation.ScriptLineNumber)"}
                    $CertUtilADTemplateOutput = $null
                    $ADTemplateOutFile = $null
                    Write-ScriptLogEntry "Error in certutil backing up $($Command.Descriptions[1]) of all certificate templates published in AD. Continuing... Error: $($LASTEXITCODE)"
                    Write-EventLogEntry "Error in certutil backing up $($Command.Descriptions[1]) of all certificate templates published in AD. Continuing... Error: $($LASTEXITCODE)" 310 "Warning"
                    continue
                }
                #Remove the last line that indicates certutil completed successfully, then output to a file
                $CertUtilADTemplateOutput.RemoveAt($($CertUtilADTemplateOutput.Count) - 1)
                try {
                    $CertUtilADTemplateOutput > $ADTemplateOutFile
                } catch {
                    Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
                    $CertUtilADTemplateOutput = $null
                    $ADTemplateOutFile = $null
                    Write-ScriptLogEntry "Error writing $($Command.Descriptions[1]) of all certificate templates published in AD to file. Continuing... Error: $($_) $($_.ScriptStackTrace)"
                    Write-EventLogEntry "Error writing $($Command.Descriptions[1]) of all certificate templates published in AD to file. Continuing... Error: $($_) $($_.ScriptStackTrace)" 311 "Warning"
                    continue
                }
            }
            if (($CertutilADTemplateOutFiles | Test-Path) -notcontains $false) {
                Write-ScriptLogEntry "Backing up a list of and the full configuration details of all certificate templates published in AD completed."
            }
            else {
                Write-ScriptLogEntry "Error in backing up one or both of a list of and the full configuration details of all certificate templates published in AD. The errors should be immediately above this entry. Continuing..."
                Write-EventLogEntry "Error in backing up one or both of a list of and the full configuration details of all certificate templates published in AD. The errors should be immediately above this entry. Continuing..." 312 "Warning"
            }
        }
        catch {
            Add-ArrayListObject $NonCriticalErrors @{"Error" = $_ ; "Location" = $_.ScriptStackTrace}
            Write-ScriptLogEntry "Error in PowerShell backing up a list of and the full configuration details of all certificate templates published in AD. Continuing... Error: $($_) $($_.ScriptStackTrace)"
            Write-EventLogEntry "Error in PowerShell backing up a list of and the full configuration details of all certificate templates published in AD. Continuing... Error: $($_) $($_.ScriptStackTrace)" 313 "Warning"
        }
    }
    else {
        Write-ScriptLogEntry "This computer is not joined to an AD domain, so backing up AD templates has been skipped. Continuing..."
    }

#This is the end of THE BIG TRY CATCH try block
}
catch {
    #Catch and store the error object of any critical error has been thrown here, so that it can be output below
    $CriticalError = $_
}

#WRAP UP
#Build and send email if email parameters were specified
#Indicate any critical or non-critical errors that occurred, or that none did, then pass to the next code block
#If any of this goes wrong in any PowerShell-terminating way, make a last-ditch effort to notify the administrator somehow
try {
    #Set-EmailNotificationContents won't do anything if -SendEmailNotification was not specified
    #It sets the values of the three variables used as parameters for Send-EmailNotification
    if ($null -ne $CriticalError) {
        Set-EmailNotificationContents "CriticalError"
        Write-ScriptLogEntry "AD CS backup script is exiting with a critical error. The error should be shown above this entry in $($LogFile). Error: $($CriticalError) $($CriticalError.ScriptStackTrace) Exiting..."
        Write-EventLogEntry "AD CS backup script is exiting with a critical error. The error should be shown both above this entry in this event log and in $($LogFile). Error: $($CriticalError) $($CriticalError.ScriptStackTrace) Exiting..." 550 "Error"
        Send-EmailNotification -Subject $EmailNotificationSubject -Body $EmailNotificationBody -Priority $EmailNotificationPriority
    }
    elseif ($($NonCriticalErrors.Count) -ge 1) {
        Set-EmailNotificationContents "NonCriticalError"
        Write-ScriptLogEntry "AD CS backup script has completed with one or more non-critical errors. See $($LogFile) or the `"AD CS Backup Script`" Windows event log for more information."
        Write-ScriptLogEntry "The following is a list of all detected non-critical errors:" -HideFromConsole
        $NonCriticalErrors | ForEach-Object { [PSCustomObject]$_ } | Format-Table -AutoSize -HideTableHeaders -Wrap | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        Write-ScriptLogEntry "If there is no output above, something went wrong with error collection. See $($LogFile) for the errors and more information." -HideFromConsole
        Write-EventLogEntry "AD CS backup script has completed. One or more non-critical errors occurred. This event log may contain those errors. If not, see $($LogFile). Exiting..." 350 "Warning"
        Send-EmailNotification -Subject $EmailNotificationSubject -Body $EmailNotificationBody -Priority $EmailNotificationPriority
    }
    else {
        Set-EmailNotificationContents "Success"
        Write-ScriptLogEntry "AD CS backup script has completed. No errors were detected. Exiting..."
        Write-EventLogEntry "AD CS backup script has completed. No errors were detected. Exiting..." 2 "Information"
        Send-EmailNotification -Subject $EmailNotificationSubject -Body $EmailNotificationBody -Priority $EmailNotificationPriority
    }
}
catch {
    Set-EmailNotificationContents "CriticalError"
    if (Test-Path -Path $LogFile | Out-Null) {
        #Try to write to the logfile if it exists
        "$(Get-TimeStamp) Unexpected critical error in AD CS backup script error handling or logging logic. Error: $($_) $($_.ScriptStackTrace) Exiting. Any gathered critical and non-critical error details below:" | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        $CriticalError | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        "---Non-Critical-Below---" | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        $NonCriticalErrors | ForEach-Object { [PSCustomObject]$_ } | Format-Table -AutoSize -HideTableHeaders -Wrap | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
    }
    elseif (Test-Path -Path $WorkingBackupPath | Out-Null) {
        #Try to write to a new file in $WorkingBackupPath if the logfile doesn't exist but the backup folder does
        $PanicLogFile = "$($WorkingBackupPath)\ADCSBACKUP-UNEXPECTED-CRITICAL-ERROR.log"
        New-Item $PanicLogFile -Force -ErrorAction SilentlyContinue
        "$(Get-TimeStamp) Unexpected critical error in AD CS backup script error handling or logging logic. Error: $($_) $($_.ScriptStackTrace) Exiting. Any gathered critical and non-critical error details below:" | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
        $CriticalError | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
        "---Non-Critical-Below---" | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        $NonCriticalErrors | ForEach-Object { [PSCustomObject]$_ } | Format-Table -AutoSize -HideTableHeaders -Wrap | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
    }
    else {
        #If all else fails, try to write to the directory the script is executing from
        $PanicLogFile = "$($MyInvocation.PSScriptRoot)\ADCSBACKUP-UNEXPECTED-CRITICAL-ERROR.log"
        New-Item $PanicLogFile -Force -ErrorAction SilentlyContinue
        "$(Get-TimeStamp) Unexpected critical error in AD CS backup script error handling or logging logic. Error: $($_) $($_.ScriptStackTrace) Exiting. Any gathered critical and non-critical error details below:" | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
        $CriticalError | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
        "---Non-Critical-Below---" | Out-File -FilePath $LogFile -Append -Force -ErrorAction SilentlyContinue
        $NonCriticalErrors | ForEach-Object { [PSCustomObject]$_ } | Format-Table -AutoSize -HideTableHeaders -Wrap | Out-File -FilePath $PanicLogFile -Append -Force -ErrorAction SilentlyContinue
    }
    Send-EmailNotification -Subject $EmailNotificationSubject -Body $EmailNotificationBody -Priority $EmailNotificationPriority
    try {
        #Throw an error and exit the script here instead of continuing to the final code block. Still clear the variables via the finally block
        throw "$(Get-TimeStamp) Unexpected critical error in AD CS backup script error handling or logging logic. Error: $($_) $($_.ScriptStackTrace) Exiting. Check $($MyInvocation.PSScriptRoot) if there is no information in $($LogFile) or $($WorkingBackupPath)."
    }
    finally {
        Get-Variable -Scope Script | Remove-Variable -Force -ErrorAction SilentlyContinue
    }
}

#If a critical error occurred, throw it and exit. Whether this happens or not, run the finally block
try {
    if ($null -ne $CriticalError) {
        throw $CriticalError
    }
}
finally {
    #Dump out all of our variables in case any of them somehow persist and the script is run again, which generally would only happen in an interactive session
    Get-Variable -Scope Script | Remove-Variable -Force -ErrorAction SilentlyContinue
}

#SCRIPT END