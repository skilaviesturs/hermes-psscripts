#Requires -Version 5.1
Function Test-VSkRemoteComputer {
    Param(
        [Parameter(Position = 0, Mandatory)]
        [string[]]$ComputerName
    )
    
    # $VerifyCompsResultFile = "$($Script:DataDir)\VerifyCompsResult.dat"
        
    Write-Host "[Test-VSkRemoteComputer] got for testing [$($ComputerName.Count)] $(
        if($ComputerName.Count -eq 1){"computer"}else{"computers"}
    )"
    Write-Verbose "[Test-VSkRemoteComputer]:got:[$($ComputerName)]"

    # nosūtam uzdevumu uz remote datoriem

    $JobResults = Send-VSkJob -Computers $ComputerName -ScriptBlockName 'SBVerifyComps'

    $JobResults | ForEach-Object {
        Write-Host "$_"
        foreach ( $line in $_.msgCatchErr ) {
            Write-Host "$line"
        }

    }

    Write-Verbose "[Test-VSkRemoteComputer] got results..."
    # Analizējam saņemtos rezultātus

    $JobResults = $JobResults | Where-Object -Property Computer -ne $null

    if ( $JobResults -ne $False -or
        $JobResults -notlike 'False' -or 
        $JobResults.Count -gt 0 ) {

        #Analizējam JObbu atgrieztos rezultātus, ievietojam DelComps
        $DelComps = @(
            Foreach ($computer in $JobResults) {
                if ( $computer.msgCatchErr.Count -gt 0 ) {
                    $computer.msgCatchErr | ForEach-Object {
                        Write-Host "$_"
                        Write-msg -log -bug -text "$_" 
                    }
                    $computer.Computer
                    Continue
                }
                if ( $computer.Version -ne '5' ) {
                    Write-msg -log -text "[Test-VSkRemoteComputer] Computer [$($computer.Computer)] `
                        Powershell version is [$($computer.Version)], must be [5]. Removed."
                    $computer.Computer
                    Continue
                }
                if ( $computer.Policy -ne 'Unrestricted' -and $computer.Policy -ne 'RemoteSigned' ) {
                    Write-msg -log -text "[Test-VSkRemoteComputer] Computer [$($computer.Computer)] powershell Execution Policy `
                        is set [$($computer.Policy)], must be [Unrestricted] or [RemoteSigned]. Removed."
                    $computer.Computer
                    Continue
                }
                if ( $computer.Language -ne 'FullLanguage' ) {
                    Write-msg -log -text "[Test-VSkRemoteComputer] Computer [$($computer.Computer)] powershell Language is set `
                        [$($computer.Language)], must be [FullLanguage]. Removed."
                    $computer.Computer
                    Continue
                }
            }
        )

        #Atbrīvojamies no dublikātiem, ja tādu ir
        $DelComps = $DelComps | Get-Unique

        #Parsējam input masīvu un papildinām DelComps ar dzēšamajiem datoriem, kas nav atbildējuši uz ping
        $ComputerName | ForEach-Object {
            if ( $JobResults.Computer.Contains($_) -eq $False ) {
                $DelComps += $_
            }
        }

        #Aizvācam no input masīva visus datorus, kas nav izturējuši pārbaudi
        $DelComps | ForEach-Object { 
            if ( $ComputerName.Contains($_) ) {
                $ComputerName = $ComputerName -ne $_
                Write-msg -log -text "[Test-VSkRemoteComputer] Computer [$_] is not ready PSRemote. Removed."
            }
        }
    }
    else {
        Write-msg -log -bug -text "[Test-VSkRemoteComputer] Oopps!! Jober returned nothing." 
    }
    if ( $JobResults.GetType().BaseType.name -eq 'Object' ) {
        $tmpJobResults = $JobResults.psobject.copy()
        $JobResults = @()
        $JobResults += @($tmpJobResults)
    }
    if ( $JobResults.count -gt 0) {
        Write-Verbose "[Test-VSkRemoteComputer]:return:[$($ComputerName)]"
        # $JobResults | Export-Clixml -Path $VerifyCompsResultFile -Depth 10 -Force
        $ComputerName
    }
    else {
        $false
    }
}
#endregion