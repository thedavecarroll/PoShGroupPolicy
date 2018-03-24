function Get-GPSetting {
<#
.SYNOPSIS
Get the group policy settings for specific extensions or list configured extensions.

.DESCRIPTION
This function will display the group policy settings of the supplied group policy
objects, or will display the extensions for which the group policy contains 
configuration settings.

It uses the GenerateReport() method of the group policy .Net object to create an
XML version of the report which it then parses.

Since generating XML can be time consuming, the Verbose common parameter can be
specified which will provide notices to the user while it is processing.

.PARAMETER GroupPolicy
The group policy or array of group policies for which to return the settings.

.PARAMETER Type
The type of data returned. Accepted values are:

ExtensionType
This type of returned data will identify which group policy extensions contain
configuration. This can return extension types beyond what the function can
currently process.

Script
DriveMapSetting
SecuritySetting
RegistrySetting
FolderRedirectionSetting

.INPUTS
Microsoft.GroupPolicy.Gpo[]

.OUTPUTS
System.Management.Automation.PSCustomObject

.EXAMPLE
C:\PS> Get-GPO -All | Get-GPSetting -Verbose -OutVariable MyGPOs

.EXAMPLE
C:\PS> Get-GPO 'DriveMapping' | Get-GPSetting -Type DriveMapping

Name                 : DriveMapping
ConfigurationGroup   : User
SettingChanged       : 2017-10-03 18:20:54
Order                : 1
DriveAction          : Replace
ShowThisDrive        : True
ShowAllDrives        : True
Label                : MyUser
Path                 : \\MyServerA\ShareB
Reconnect            : True
FirstAvailableLetter : False
DriveLetter          : P
ConnectUserName      :
Filters              : CONTOSO\MyUsers

Name                 : DriveMapping
ConfigurationGroup   : User
SettingChanged       : 2017-10-03 18:20:57
Order                : 2
DriveAction          : Replace
ShowThisDrive        : True
ShowAllDrives        : True
Label                : Department
Path                 : \\DeptServer\ShareB
Reconnect            : True
FirstAvailableLetter : False
DriveLetter          : S
ConnectUserName      :
Filters              : CONTOSO\DeptUsers

.EXAMPLE
C:\PS> Get-GPO 'Workstation Scripts' | Get-GPSetting -Type Script | Sort-Object -Property Type,Order | Format-Table -AutoSize

Name                ConfigurationGroup Script                   Type     Parameters Order PSRunOrder
----                ------------------ ------                   ----     ---------- ----- ----------
Workstation Scripts Computer           CleanTempFiles.cmd       Shutdown            0     PSNotConfigured
Workstation Scripts Computer           ComputerInventory.ps1    Startup             0     RunPSFirst
Workstation Scripts Computer           InstallApps.cmd          Startup             1     RunPSFirst
Workstation Scripts Computer           ClearAppCache.vbs        Startup             2     RunPSFirst

.LINK
https://github.com/thedavecarroll/PoShGroupPolicy

#>
    [CmdLetBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline)]
        [Microsoft.GroupPolicy.Gpo[]]$GroupPolicy,
        
        [ValidateSet('ExtensionType','Script','DriveMapSetting','SecuritySetting','RegistrySetting','FolderRedirectionSetting')]
        [String]$Type='ExtensionType'
    )

    begin {
        $Timer = [System.Diagnostics.Stopwatch]::StartNew()
        $Counter = 0
    }

    process {

        foreach ($GPO in $GroupPolicy) {

            Write-Verbose -Message "Generating XML report to parse $Type data for GPO $($GPO.DisplayName)"
            [xml]$GpoXml = $GPO.GenerateReport('Xml')

            $GpoName      = $GpoXml.GPO.Name
            $CreatedTime  = $GpoXml.GPO.CreatedTime
            $ModifiedTime = $GpoXml.GPO.ModifiedTime
            $ReadTime     = $GpoXml.GPO.ReadTime

            $Counter++
            if ($VerbosePreference -eq 'Continue') {
                Write-Progress -Activity "Processing..." -CurrentOperation $GpoName -Status "$Counter / $($GroupPolicy.Count)"
            }

            switch ($Type) {
                'ExtensionType' {
                    foreach ($Extension in $GpoXml.GPO.Computer.ExtensionData.Extension) {
                        [PsCustomObject]@{
                            Name = $GpoName
                            ConfigurationGroup = 'Computer'
                            CreatedTime = $CreatedTime
                            ModifiedTime = $ModifiedTime
                            ExtensionType = $Extension.Type.Split(":")[1]
                        }
                    }
                    foreach ($Extension in $GpoXml.GPO.User.ExtensionData.Extension) {
                        [PsCustomObject]@{
                            Name = $GpoName
                            ConfigurationGroup = 'User'
                            CreatedTime = $CreatedTime
                            ModifiedTime = $ModifiedTime
                            ExtensionType = $Extension.Type.Split(":")[1]
                        }
                    }

                }
                'Script' {
                    foreach ($Script in $Gpoxml.GPO.Computer.ExtensionData.Extension.Script ) {
                        if ($Script) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'Computer'
                                Script = $Script.Command
                                Type = $Script.Type
                                Parameters = $Script.Parameters
                                Order = $Script.Order
                                PSRunOrder = $Script.RunOrder
                            }
                        }
                    }
                    foreach ($Script in $Gpoxml.GPO.User.ExtensionData.Extension.Script ) {
                        if ($Script) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'User'
                                Script = $Script.Command
                                Type = $Script.Type
                                Parameters = $Script.Parameters
                                Order = $Script.Order
                                PSRunOrder = $Script.RunOrder
                            }
                        }
                    }
                }
                'DriveMapSetting' {
                    foreach ($DriveMapping in $Gpoxml.GPO.Computer.ExtensionData.Extension.DriveMapSettings.Drive ) {
                        if ($DriveMapping) {
                            switch ($DriveMapping.Properties.action) {
                                'R' { $DriveAction = 'Replace'}
                                'U' { $DriveAction = 'Update'}
                                'C' { $DriveAction = 'Create'}
                                'D' { $DriveAction = 'Delete'}
                            }
                            if ($DriveMapping.Properties.persistent -eq 1) {
                                $Reconnect = $true
                            } else {
                                $Reconnect = $false
                            }
                            if ($DriveMapping.Filters) {
                                $Filters = $DriveMapping.Filters.FilterGroup.Name
                            } else {
                                $Filters = $null
                            }
                            if ($DriveMapping.Properties.thisDrive -eq 'SHOW') {
                                $ShowThisDrive = $true
                            } elseif ($DriveMapping.Properties.thisDrive -eq 'HIDE') {
                                $ShowThisDrive = $false
                            } else {
                                $ShowThisDrive = $null
                            }
                            if ($DriveMapping.Properties.allDrives -eq 'SHOW') {
                                $ShowAllDrives = $true
                            } elseif ($DriveMapping.Properties.allDrives -eq 'HIDE') {
                                $ShowAllDrives = $false
                            } else {
                                $ShowAllDrives = $null
                            }
                            if ($DriveMapping.Properties.useLetter -eq 0) {
                                $FirstAvailableLetter = $true
                            } else {
                                $FirstAvailableLetter = $false
                            }
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'Computer'
                                SettingChanged = $DriveMapping.Changed
                                Order = $DriveMapping.GPOSettingOrder                                
                                DriveAction = $DriveAction
                                ShowThisDrive = $ShowThisDrive
                                ShowAllDrives = $ShowAllDrives                                
                                Label = $DriveMapping.Properties.label
                                Path = $DriveMapping.Properties.path
                                Reconnect = $Reconnect
                                FirstAvailableLetter = $FirstAvailableLetter
                                DriveLetter = $DriveMapping.Properties.letter
                                ConnectUserName = $DriveMapping.Properties.userName
                                Filters = $Filters
                            }
                        }
                    }
                    foreach ($DriveMapping in $Gpoxml.GPO.User.ExtensionData.Extension.DriveMapSettings.Drive ) {
                        if ($DriveMapping) {
                            switch ($DriveMapping.Properties.action) {
                                'R' { $DriveAction = 'Replace'}
                                'U' { $DriveAction = 'Update'}
                                'C' { $DriveAction = 'Create'}
                                'D' { $DriveAction = 'Delete'}
                            }
                            if ($DriveMapping.Properties.persistent -eq 1) {
                                $Reconnect = $true
                            } else {
                                $Reconnect = $false
                            }
                            if ($DriveMapping.Filters) {
                                $Filters = $DriveMapping.Filters.FilterGroup.Name
                            } else {
                                $Filters = $null
                            }
                            if ($DriveMapping.Properties.thisDrive -eq 'SHOW') {
                                $ShowThisDrive = $true
                            } elseif ($DriveMapping.Properties.thisDrive -eq 'HIDE') {
                                $ShowThisDrive = $false
                            } else {
                                $ShowThisDrive = $null
                            }
                            if ($DriveMapping.Properties.allDrives -eq 'SHOW') {
                                $ShowAllDrives = $true
                            } elseif ($DriveMapping.Properties.allDrives -eq 'HIDE') {
                                $ShowAllDrives = $false
                            } else {
                                $ShowAllDrives = $null
                            }
                            if ($DriveMapping.Properties.useLetter -eq 0) {
                                $FirstAvailableLetter = $true
                            } else {
                                $FirstAvailableLetter = $false
                            }
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'User'
                                SettingChanged = $DriveMapping.Changed
                                Order = $DriveMapping.GPOSettingOrder                                
                                DriveAction = $DriveAction
                                ShowThisDrive = $ShowThisDrive
                                ShowAllDrives = $ShowAllDrives                                
                                Label = $DriveMapping.Properties.label
                                Path = $DriveMapping.Properties.path
                                Reconnect = $Reconnect
                                FirstAvailableLetter = $FirstAvailableLetter
                                DriveLetter = $DriveMapping.Properties.letter
                                ConnectUserName = $DriveMapping.Properties.userName
                                Filters = $Filters
                            }
                        }
                    }
                }
                'SecuritySetting' {
                    foreach ($SecuritySetting in $Gpoxml.GPO.Computer.ExtensionData.Extension.SecurityOptions ) {
                        if ($SecuritySetting) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'Computer'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                KeyName = $SecuritySetting.KeyName
                                SettingNumber = $SecuritySetting.SettingNumber
                                Display = $SecuritySetting.Display.Name
                                Units = $SecuritySetting.Display.Units
                            }
                        }
                    }
                    foreach ($SecuritySetting in $Gpoxml.GPO.User.ExtensionData.Extension.SecurityOptions ) {
                        if ($SecuritySetting) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'User'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                KeyName = $SecuritySetting.KeyName
                                SettingNumber = $SecuritySetting.SettingNumber
                                Display = $SecuritySetting.Display.Name
                                Units = $SecuritySetting.Display.Units
                            }
                        }
                    }
                }
                'RegistrySetting' {
                    foreach ($RegistrySetting in $Gpoxml.Computer.ExtensionData.Extension.Policy ) {
                        if ($RegistrySetting) {
                            $GPORegistrySettingsInfo += [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'Computer'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                PolicyName = $RegistrySetting.Name
                                State = $RegistrySetting.State
                                Supported = $RegistrySetting.Supported
                            }
                        }
                    }
                    foreach ($RegistrySetting in $Gpoxml.User.ExtensionData.Extension.Policy ) {
                        if ($RegistrySetting) {
                            $GPORegistrySettingsInfo += [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'User'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                PolicyName = $RegistrySetting.Name
                                State = $RegistrySetting.State
                                Supported = $RegistrySetting.Supported
                            }
                        }
                    }
                }
                'FolderRedirectionSetting' {
                    try {
                        $KnownFolders = Get-GPKnownFolderId -ErrorAction Stop                 
                        foreach ($FolderRedirectionSetting in $Gpoxml.Computer.ExtensionData.Extension.Folder ) {
                            if ($FolderRedirectionSetting) {
                                $GPOFolderRedirectionSettingsInfo += [PsCustomObject]@{
                                    Name = $GpoName
                                    ConfigurationGroup = 'Computer'
                                    CreatedTime = $CreatedTime
                                    ModifiedTime = $ModifiedTime
                                    Id = $FolderRedirectionSetting.Id
                                    DisplayName = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty DisplayName
                                    FolderType = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty FolderType
                                    DefaultPath = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty DefaultPath
                                    DestinationPath = $FolderRedirectionSetting.Location.DestinationPath
                                }
                            }
                        }
                        foreach ($FolderRedirectionSetting in $Gpoxml.User.ExtensionData.Extension.Folder ) {
                            if ($FolderRedirectionSetting) {
                                $GPOFolderRedirectionSettingsInfo += [PsCustomObject]@{
                                    Name = $GpoName
                                    ConfigurationGroup = 'User'
                                    CreatedTime = $CreatedTime
                                    ModifiedTime = $ModifiedTime
                                    Id = $FolderRedirectionSetting.Id
                                    DisplayName = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty DisplayName
                                    FolderType = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty FolderType
                                    DefaultPath = $KnownFolders | Where-Object {$_.GUID -eq $FolderRedirectionSetting.Id} | Select-Object -ExpandProperty DefaultPath
                                    DestinationPath = $FolderRedirectionSetting.Location.DestinationPath
                                }
                            }
                        }
                    }
                    catch {
                        Write-Warning -Message 'Unable to obtain list of KnownFolders.'
                    }
                }
            } # end switch

        } # end foreach loop

    }

    end {
        Write-Verbose -Message "Completed in $([system.String]::Format("{0}d {1:00}h:{2:00}m:{3:00}s.{4:00}", $Timer.Elapsed.Days, $Timer.Elapsed.Hours, $Timer.Elapsed.Minutes, $Timer.Elapsed.Seconds, $Timer.Elapsed.Milliseconds / 10))"
    }
}