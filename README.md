# ClearOutlookCache

There are two scripts in this repository that fulfil the same task.


Script 1 *ClearOutlookCache_M.ps1* is a PowerShell script designed for manual execution. This script can be executed on the respective device by the user or administrator.

Script 2 *ClearOutlookCache_A.ps1* mirrors the functionality of Script 1 but is tailored for deployment via Microsoft Intune. 

The differences between the two scripts relate to execution and logging.

Log files are created for both scripts, these are stored under C:\MDM\Logging.

With the ClearOutlookCache_A.ps1 script, which is optimised for Intune, a REG key is created after execution. This prevents the script from being executed at every restart.

If the script has to be executed again, the following REG key must be deleted
HKCU:Software\Microsoft\MSB365_Outlook_clear_cache_Tool

The tasks that are automated by these scripts can also be executed manually. More information can be found under the following link:

https://techgenix.com/creating-a-new-outlook-profile-without-user-involvement/


More information about the Intune configuration, can be found at: https://www.msb365.blog/?p=5479
