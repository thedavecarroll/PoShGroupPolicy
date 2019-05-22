function Get-RsGpoReport {
    [CmdLetBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline)]
        [Microsoft.GroupPolicy.Gpo[]]$GroupPolicy,

        [int]$MaxThreads = 10
    )

    begin {
        $WorkerPool = [System.Collections.Generic.List[Object]]::new()

        $Invokes = @()
        $Instances = @()

        $SessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        $RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1,$MaxThreads)
        $PowerShell = [System.Management.Automation.PowerShell]::Create()
        $PowerShell.RunspacePool = $RunspacePool
        $RunspacePool.Open()
    }

    process {

        foreach ($GPO in $GroupPolicy) {

            $Worker = [System.Management.Automation.PowerShell]::Create()
            $Worker.RunspacePool = $RunspacePool

            $ScriptBlock = {
                param ($GPO)
                    $GPO.DisplayName | Write-Verbose -Verbose

                    $ManagedThreadId = [System.Threading.Thread]::CurrentThread.ManagedThreadId
                    'Managed Thread Id: Beginning {0}' -f $ManagedThreadId | Write-Verbose -Verbose

                    try {
                        'Generate XML Report...' | Write-Verbose -Verbose
                        [void]$GPO.GenerateReport('xml')
                        'Generate XML Report...completed' | Write-Verbose -Verbose
                    }
                    catch {
                        $_
                    }
                    'Managed Thread Id: Ending {0}' -f $ManagedThreadId | Write-Verbose -Verbose
            }

            [void]$Worker.AddScript($ScriptBlock).AddArgument($GPO)

            $Handle = $Worker.BeginInvoke()
            $WorkerPool.Add(
                [PSCustomObject]@{
                    Worker = $Worker
                    Handle = $Handle
                }
            )

            'Available Runspaces in RunspacePool: {0}' -f $RunspacePool.GetAvailableRunspaces() | Write-Debug
            'Remaining Jobs: {0}' -f @($WorkerPool | Where-Object { $_.Handle.iscompleted -ne 'Completed'}).Count | Write-Debug
        }

    }

    end {

        $WorkerPool | ForEach-Object {
            $_.Worker.EndInvoke($_.Handle)
            $_.Worker.Dispose()
        }

        $RunspacePool.Close()
    }

}