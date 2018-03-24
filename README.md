# PoShGroupPolicy 0.3
PowerShell module to assist with Group Policy

## Get the module
PoShGroupPolicy can be downloaded or insptected at the [PowerShellGallery](https://www.powershellgallery.com/packages/PoShGroupPolicy)

## What GPO does that again?
Have you ever need to update a script that you knew was configured in a group policy, but you just didn't know which 
GPO? This module helps the by parsing the group policy and returning key pieces of data for several GP extension types.

## Example
```powershell
C:\PS> Get-GPO 'Workstation Scripts' | Get-GPSetting -Type Script | Sort-Object -Property Type,Order | Format-Table -AutoSize

Name                ConfigurationGroup Script                   Type     Parameters Order PSRunOrder
----                ------------------ ------                   ----     ---------- ----- ----------
Workstation Scripts Computer           CleanTempFiles.cmd       Shutdown            0     PSNotConfigured
Workstation Scripts Computer           ComputerInventory.ps1    Startup             0     RunPSFirst
Workstation Scripts Computer           InstallApps.cmd          Startup             1     RunPSFirst
Workstation Scripts Computer           ClearAppCache.vbs        Startup             2     RunPSFirst
```

### Note
This is my first public module on GitHub and I'm eager to learn. Feel free to submit suggestions and especially any
corrections. Please watch or Find-Module -Name PoshGroupPolicy occassionally to see if I've published any updates.
