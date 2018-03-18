function Get-GPSetting {
    [CmdLetBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [Microsoft.GroupPolicy.Gpo[]]$GroupPolicy,
        
        [ValidateSet('ExtensionType','Script','DriveMapSetting','SecuritySetting','RegistrySetting','FolderRedirectionSetting')]
        [String]$Type='ExtensionType'
    )

    begin {
        $Timer = [System.Diagnostics.Stopwatch]::StartNew()
    }

    process {

        foreach ($GPO in $GroupPolicy) {

            Write-Verbose -Message "Generating XML report to parse $Type data for GPO $($GPO.DisplayName)"
            [xml]$GpoXml = $GPO.GenerateReport('Xml')

            $GpoName      = $GpoXml.GPO.Name
            $CreatedTime  = $GpoXml.GPO.CreatedTime
            $ModifiedTime = $GpoXml.GPO.ModifiedTime

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
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
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
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
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
                    foreach ($DriveMapping in $Gpoxml.GPO.Computer.ExtensionData.Extension.DriveMapSettings.Drive.properties ) {
                        if ($DriveMapping) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'Computer'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                DriveAction = $DriveMapping.action
                                ThisDrive = $DriveMapping.thisDrive
                                AllDrives = $DriveMapping.allDrives
                                UserName = $DriveMapping.userName
                                Path = $DriveMapping.path
                                Persistent = $DriveMapping.persistent
                                UseLetter = $DriveMapping.useLetter
                                DriveLetter = $DriveMapping.letter
                            }
                        }
                    }
                    foreach ($DriveMapping in $Gpoxml.GPO.User.ExtensionData.Extension.DriveMapSettings.Drive.properties ) {
                        if ($DriveMapping) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                ConfigurationGroup = 'User'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                DriveAction = $DriveMapping.action
                                ThisDrive = $DriveMapping.thisDrive
                                AllDrives = $DriveMapping.allDrives
                                UserName = $DriveMapping.userName
                                Path = $DriveMapping.path
                                Persistent = $DriveMapping.persistent
                                UseLetter = $DriveMapping.useLetter
                                DriveLetter = $DriveMapping.letter
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