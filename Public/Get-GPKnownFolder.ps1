function Get-GPKnownFolder {
    $WebRequest = Invoke-WebRequest -Uri "https://msdn.microsoft.com/en-us/library/windows/desktop/dd378457(v=vs.85).aspx"
    $Ids = $WebRequest.AllElements | Where-Object {$_.Class -eq "mtps-table clsStd"} | Select-Object -ExpandProperty InnerText

    foreach ($Record in $Ids) {
        [PsCustomObject]@{
            GUID = $Record.Split("`n")[0].Replace("GUID","").Trim()
            DisplayName = $Record.Split("`n")[1].Replace("Display Name","").Trim()
            FolderType = $Record.Split("`n")[2].Replace("Folder Type","").Trim()
            DefaultPath = $Record.Split("`n")[3].Replace("Default Path","").Trim()
            CSIDLEquivalent = $Record.Split("`n")[4].Replace("CSIDL Equivalent","").Trim()
            LegacyDisplayName = $Record.Split("`n")[5].Replace("Legacy Display Name","").Trim()
            LegacyDefaultPath = $Record.Split("`n")[6].Replace("Legacy Default Path","").Trim()
        }
    }
}