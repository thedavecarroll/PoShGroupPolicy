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
        
        #[ValidateSet('ExtensionType','Script','DriveMapSetting','SecuritySetting','RegistrySetting','FolderRedirectionSetting')]
        [ValidateNotNullOrEmpty()]
        [String]$Type#='ExtensionType'
    )

    begin {

        Write-Verbose -Message 'Importing Group Policy module...'
        try {
            Import-Module -Name GroupPolicy -Verbose:$false -ErrorAction stop
        }
        catch {
            Write-Warning -Message 'Failed to import GroupPolicy module'
            exit 1
        }

        $Timer = [System.Diagnostics.Stopwatch]::StartNew()
        $Counter = 0
        $TotalCount = $GroupPolicy.Count
        
    }

    process {

        foreach ($GPO in $GroupPolicy) {

            $Counter++
            if ($VerbosePreference -eq 'Continue' -and $TotalCount -gt 1) {
                Write-Verbose -Message "Generating XML report to parse GPO $($GPO.DisplayName)"
                Write-Progress -Activity "Processing..." -CurrentOperation $GPO.DisplayName -Status "$Counter / $TotalCount"
            }                
            
            try {
                [xml]$GpoXml = $GPO.GenerateReport('Xml')
            }
            catch {
                Write-Warning -Message 'Unable to generate XML report'
                Write-Warning -Message $_.Exception.Message
                continue            
            }

            $XmlNamespaceManager = [System.Xml.XmlNamespaceManager]::New($GpoXml.CreateNavigator().NameTable)
            $XmlNamespaces = $GpoXml.CreateNavigator().GetNamespacesInScope('All')
            foreach ($key in $XmlNamespaces.keys) { 
                $XmlNamespaceManager.AddNamespace( $key, $XmlNamespaces.$key ) 
            }
            $XmlNamespaceManager.AddNamespace('gp','http://www.microsoft.com/GroupPolicy/Settings')

            #$ExtensionNodes = $GpoXml.SelectNodes("/gp:GPO//gp:ExtensionData[gp:Name = '$Type']/gp:Extension", $XmlNamespaceManager)
            $ExtensionNodes = $GpoXml.SelectNodes("/gp:GPO//gp:ExtensionData/gp:Extension", $XmlNamespaceManager)
            $ExtensionNodes = $GpoXml.SelectNodes("//gp:ExtensionData/gp:Extension", $XmlNamespaceManager)

            $GPConfiguration = foreach ($Node in $ExtensionNodes.ChildNodes) { 
                
                $Properties = $Node | Get-Member -MemberType Property | Select-Object -ExpandProperty Name -Unique
                $Settings = $Node | Select-Object -Property $Properties                    
                
                $TypeName = (($Node.NamespaceURI -Split '/' | Select-Object -Skip 3) -Join '.') + '.' + $Node.LocalName

                foreach ($Setting in $Settings) {
                    
                    $GPSettings = [PsCustomObject]@{    
                        GpoName = $Node.ParentNode.ParentNode.ParentNode.ParentNode.Name
                        Guid = $Node.ParentNode.ParentNode.ParentNode.ParentNode.Identifier.Identifier.InnerText
                        DomainName  = $Node.ParentNode.ParentNode.ParentNode.ParentNode.Identifier.Domain.InnerText
                        CreatedTime  = $Node.ParentNode.ParentNode.ParentNode.ParentNode.CreatedTime
                        ModifiedTime = $Node.ParentNode.ParentNode.ParentNode.ParentNode.ModifiedTime
                        ReadTime = $Node.ParentNode.ParentNode.ParentNode.ParentNode.ReadTime
                        #ComputerConfiguration = if ($Node.ParentNode.ParentNode.ParentNode.ParentNode.Computer.Enabled) { $true } else { $false }
                        #UserConfiguration = if ($Node.ParentNode.ParentNode.ParentNode.ParentNode.User.Enabled) { $true } else { $false }
                        #LinksTo = $Node.ParentNode.ParentNode.ParentNode.ParentNode. LinksTo
                        Configuration = $Node.ParentNode.ParentNode.ParentNode.Name
                        ExtensionType  = $Node.NamespaceURI.Split('/')[-1]
                        XmlNamespace = $Node.NamespaceURI
                        TypeName = $TypeName
                        LocalName = $Node.LocalName
                        Node = $Node
                    }
                    
                    $GPSettings.PsObject.TypeNames.Insert(0,$TypeName)

                    foreach ($Property in $Properties) {
                        if ($Property -eq 'Member') {
                            $Members = @()
                            foreach ($Member in $Node.Member.ChildNodes) {
                                if ($Member.PreviousSibling -eq $null -and $Member.NextSibling -ne $null) {                                    
                                    $Members += [PsCustomObject]@{
                                        #$($Member.Name) = $Member.InnerText
                                        #$($Member.NextSibling.Name) = $Member.NextSibling.InnerText                                        
                                        (Get-Culture).TextInfo.totitlecase($Member.InnerText.ToLower()).Replace(' ','') = $Member.NextSibling.InnerText
                                    }
                                }
                            }
                        
                            Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name Members -Value $Members -Force
                        }
                        
                        switch ($Property) {
                            'Name' {
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name 'PolicyName' -Value $Setting.$Property -Force
                            }
                            'EditText' {                                
                                $EditTextProperties = $Setting.$Property | Get-Member -MemberType Property | Select-Object -ExpandProperty Name -Unique
                                $Policies = $Setting.$Property | Select-Object -Property $EditTextProperties
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name 'Settings' -Value $Policies -Force
                            }
                            'Checkbox' {
                                $CheckboxProperties = $Setting.$Property | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name -Unique
                                $Policies = $Setting.$Property | Select-Object -Property $CheckboxProperties
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name 'CheckboxSettings' -Value $Policies -Force
                            }
                            'Numeric' {
                                $NumericProperties = $Setting.$Property | Get-Member -MemberType Property | Select-Object -ExpandProperty Name -Unique
                                $Policies = $Setting.$Property | Select-Object -Property $NumericProperties
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name 'NumericSettings' -Value $Policies -Force
                            }
                            'DropDownList' {
                                $NumericProperties = $Setting.$Property | Get-Member -MemberType Property | Select-Object -ExpandProperty Name -Unique
                                $Policies = $Setting.$Property | Select-Object -Property $NumericProperties
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name 'DropDownLists' -Value $Policies -Force
                            }                            
                            'Member' {

                            }
                            default {
                                Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name $Property -Value $Setting.$Property -Force
                            }
                        }
                    }

                    $Value = $Node.ChildNodes.ChildNodes.ChildNodes | Where-Object { $_.'#text' } | Select-Object -ExpandProperty '#text'
                    if ($Value) {
                        Add-Member -InputObject $GPSettings -MemberType NoteProperty -Name Value -Value $Value -Force
                    }

                }
                
                $GPConfiguration
                
            } # end foreach node loop

        } # end foreach GPO loop

    }

    end {
        Write-Verbose -Message "Successfully processed $($GPO.count) Group Policies."
        Write-Verbose -Message "Completed in $([system.String]::Format("{0}d {1:00}h:{2:00}m:{3:00}s.{4:00}", $Timer.Elapsed.Days, $Timer.Elapsed.Hours, $Timer.Elapsed.Minutes, $Timer.Elapsed.Seconds, $Timer.Elapsed.Milliseconds / 10))"
    }
}


<#

               
 

    
                switch ($Type) {
                    'ExtensionType' {
                        foreach ($Extension in $GpoXml.GPO.Computer.ExtensionData.Extension) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                DomainName = $GPO.DomainName
                                ConfigurationGroup = 'Computer'
                                CreatedTime = $CreatedTime
                                ModifiedTime = $ModifiedTime
                                ExtensionType = $Extension.Type.Split(":")[1]
                            }
                        }
                        foreach ($Extension in $GpoXml.GPO.User.ExtensionData.Extension) {
                            [PsCustomObject]@{
                                Name = $GpoName
                                DomainName = $GPO.DomainName
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
            
    $xmlnsGpSettings = 'http://www.microsoft.com/GroupPolicy/Settings'
    $xmlnsSchemaInstance = 'http://www.w3.org/2001/XMLSchema-instance'
    $xmlnsSchema = 'http://www.w3.org/2001/XMLSchema'
    $ComputerConfiguration = 'gp:Computer/gp:ExtensionData/gp:Extension'
    $UserConfiguration = 'gp:User/gp:ExtensionData/gp:Extension'


    function Get-XmlNodeData {
        param ($ExtensionNodes)




    }

    $Drives = foreach ($node in $extensionNodes.ChildNodes) { $Props = $node | Get-Member -MemberType Property | Select-Object -ExpandProperty Name -Unique ; $node | select -Property $Props}
    PS C:\PowerShell\Temp> $Drives[0]
    
    clsid                                  Drive
    -----                                  -----
    {8FDDCC1A-0C3C-43cd-A6B4-71A6DF20DA8C} {T:, P:}
    
    
    PS C:\PowerShell\Temp> $Drives.Drive
                    
    $extensionNodes[0].ChildNodes
    foreach ($node in $extensionNodes.ChildNodes) { $Node | Get-Member -MemberType Property | Select-Object -ExpandProperty Name }
    #Display
    #KeyName
    #SettingNumber
                                
    (Get-Culture).textinfo.totitlecase('Path to theme file'.tolower()).Replace(' ','')


#>