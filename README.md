# Invoke-ADCSBackup.ps1

I wrote a backup script for Active Directory Certificate Services, and just wanted to share it in case it would be useful to anyone else.
I am not an expert in PowerShell by any means, but I am always learning.

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
-The CA's public certificate & CRL files, if in the default $env:systemroot\System32\CertSrv\CertEnroll             #Should be available on a separate CDP/AIA server
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

v1.0 - February 4, 2025
Basic debugging and testing is complete - not all failure paths tested
More optimization/best practice refinements
Certutil export steps are now loops instead of repeating the same code 
Script is fully working with no errors in a lab environment
Ready to deploy the script to my production environment

v0.95 - January 30, 2025
A lot of optimization was done - testing now underway

v0.9 - January 22, 2025
Initial "final" version of script before testing

TO DO
-Parameter validation for -BackupPath
-Separate script to copy the database backup and validate it (already have code.) You can't restore a validated database file. Go figure.

WISHLIST
-Replace CA registry backup code with PowerShell code instead of invoking reg.exe. Looks like it might not be difficult to do with the Win32 API
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

2) A batch file provided to me by Microsoft/JDA Partners Technical Service,
itself adapted from a CC BY SA 4.0 licensed batch file from Sentella.info (which is seemingly defunct in January 2025 as I cannot find it):

#Script header:
BM_DR_CABackup_v3.txt
:: Sentella BareMetal_DisasterRecovery_CA Backup Utility with Timings - Blame: GTKT 
:: This script is Modified and ©2022 Sentella.info to build and keep copy of the CA configuration
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
