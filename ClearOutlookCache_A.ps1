#region Description
<#     
       .NOTES
       ==============================================================================
       Created on:         2023/12/01 
       Created by:         Drago Petrovic
       Organization:       MSB365.blog
       Filename:           ClearOutlookCache_A.ps1
       Current version:    V1.00     

       Find us on:
             * Website:         https://www.msb365.blog
             * Technet:         https://social.technet.microsoft.com/Profile/MSB365
             * LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
             * MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
       ==============================================================================

       .DESCRIPTION
       This script is optimated for running with Microsoft Intune.
                  
       

       .NOTES
       If you need to re run this Script, you need to delete the following REG KEY: 
       HKCU:Software\Microsoft\MSB365_Outlook_clear_cache_Tool 
       CacheCleared





       .EXAMPLE
       .\ClearOutlookCache_A.ps1 
             

       .COPYRIGHT
       Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
       to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
       and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

       The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

       THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
       FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
       WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
       ===========================================================================
       .CHANGE LOG
             V1.00, 2023/12/01 - DrPe - Initial version


             
			 




--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
##############################################################################################################
IF ((Test-Path 'HKCU:Software\Microsoft\MSB365_Outlook_clear_cache_Tool') -eq $false)
{
################################################
$directory0 = "C:\MDM\Logging"
Write-Host "Checking if $directory0 is available..." -ForegroundColor Magenta
Start-Sleep -s 1
If ((Test-Path -Path $directory0) -eq $false)
{
        Write-Host "The Directory $directory0 don't exists!" -ForegroundColor Cyan
        Start-Sleep -s 2
        Write-Host "Creating directory $directory0 ..." -ForegroundColor Cyan
        Start-Sleep -s 2
        New-Item -Path $directory0 -ItemType directory
        Start-Sleep -s 2
        Write-Host "New Directory $directory0 is is created" -ForegroundColor Green
        Start-Sleep -s 2
}else{
Write-Host "The Path $directory0 already exists" -ForegroundColor green
Start-Sleep -s 3
}

$currentDateTime = Get-Date -Format "yyyyMMdd_HHmm"
$newFileName = "ClearOutlookCach_$currentDateTime.txt"
New-Item -ItemType File -Path $newFileName

################################################
Start-Transcript -Path "C:\MDM\Logging\$newFileName" -NoClobber
################################################
$directory1 = "C:\MDM\ClearOutlookCache"
Write-Host "Checking if $directory1 is available..." -ForegroundColor Magenta
Start-Sleep -s 1
If ((Test-Path -Path $directory1) -eq $false)
{
        Write-Host "The Directory $directory1 don't exists!" -ForegroundColor Cyan
        Start-Sleep -s 2
        Write-Host "Creating directory $directory1 ..." -ForegroundColor Cyan
        Start-Sleep -s 2
        New-Item -Path $directory1 -ItemType directory
        Start-Sleep -s 2
        Write-Host "New Directory $directory1 is is created" -ForegroundColor Green
        Start-Sleep -s 2
}else{
Write-Host "The Path $directory1 already exists" -ForegroundColor green
Start-Sleep -s 3
}

################################################################################################

$registryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook"

if (Test-Path $registryPath) {
    Write-Host "The registry key $registryPath exists." -ForegroundColor Green
    Remove-Item -Path $registryPath -Recurse -Force -Confirm:$false
} else {
    Write-Host "The registry key $registryPath does not exist." -ForegroundColor DarkMagenta -BackgroundColor Gray
}

################################################################################################
################################################
New-Item -Path HKCU:Software\Microsoft\MSB365_Outlook_clear_cache_Tool
Get-Item -Path "HKCU:Software\Microsoft\MSB365_Outlook_clear_cache_Tool" | New-ItemProperty -Name CacheCleared -Value yes
################################################

}else{
$directory0 = "C:\MDM\Logging"
Write-Host "Checking if $directory0 is available..." -ForegroundColor Magenta
Start-Sleep -s 1
If ((Test-Path -Path $directory0) -eq $false)
{
        Write-Host "The Directory $directory0 don't exists!" -ForegroundColor Cyan
        Start-Sleep -s 2
        Write-Host "Creating directory $directory0 ..." -ForegroundColor Cyan
        Start-Sleep -s 2
        New-Item -Path $directory0 -ItemType directory
        Start-Sleep -s 2
        Write-Host "New Directory $directory0 is is created" -ForegroundColor Green
        Start-Sleep -s 2
}else{
Write-Host "The Path $directory0 already exists" -ForegroundColor green
Start-Sleep -s 3
}
################################################
Start-Transcript -Path "C:\MDM\Logging\ClearOutlookCach.txt" -NoClobber
################################################
Write-Host "No action needed. The Outlook cache has already been cleared!" -ForegroundColor Green
Stop-Transcript
}
