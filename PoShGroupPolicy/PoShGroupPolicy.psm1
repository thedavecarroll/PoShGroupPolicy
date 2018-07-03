[CmdLetBinding()]
param()

#Requires -Modules GroupPolicy

# dot source public and private function definition files, export publich functions
try {
    foreach ($Scope in 'Public','Private') {
        Get-ChildItem "$PSScriptRoot\$Scope" -Filter *.ps1 | ForEach-Object {
            . $_.FullName
            if ($Scope -eq 'Public') { 
                Export-ModuleMember -Function $_.BaseName -ErrorAction Stop
            }            
        }
    }
} 
catch {
    Write-Error ("{0}: {1}" -f $_.BaseName,$_.Exception.Message)
    exit 1
}

# import format and type data
Try {
    Update-FormatData "$PSScriptRoot\TypeData\PoshGroupPolicy.Format.ps1xml" -ErrorAction Stop
}
catch {
    Write-Error ("{0}: {1}" -f 'Update-FormatData',$_.Exception.Message)
    exit 1
}
#try {
    #Update-TypeData "$PSScriptRoot\TypeData\PoshRSJob.Types.ps1xml" -ErrorAction Stop
#}
#catch {}
