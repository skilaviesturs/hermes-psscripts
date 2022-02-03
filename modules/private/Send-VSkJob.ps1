#Requires -Version 5.1
Function Send-VSkJob {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory)]
        [string[]]$Computers,
        [Parameter(Position = 1, Mandatory)]
        [string]$ScriptBlockName,
        [switch]$Update,
        [switch]$AutoReboot,
        [System.IO.FileInfo]$InstallPath,
        [string]$DataArchiveFile,
        [string]$UninstallEncryptedParameter
    )
    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
    }
    $MaxJobsThreads = 25
    Write-Verbose "[Send-VSkJob]:ComputerName[$Computers]; ScriptBlockName[$ScriptBlockName]"
    Write-Verbose "[Send-VSkJob]:Update [$Update]; AutoReboot[$AutoReboot]"
    Write-Verbose "[Send-VSkJob]:InstallPath [$($InstallPath.FullName)]"
    Write-Verbose "[Send-VSkJob]:DataArchiveFile [$DataArchiveFile]"
    Write-Verbose "[Send-VSkJob]:UninstallEncryptedParameter [$UninstallEncryptedParameter]"

    $jobWatch = [System.Diagnostics.Stopwatch]::startNew()
    $Output = @()
    Write-Host -NoNewLine "[Set-VSkJobs] Running jobs : " -ForegroundColor Yellow -BackgroundColor Black

    ForEach ( $Computer in $Computers ) {
        While ($(Get-Job -State "Running").count -ge $MaxJobsThreads) {
            Start-Sleep -Milliseconds 10
        }
        
        if ( $ScriptBlockName -eq 'Get-VSkRemoteComputerInfo' ) { 
            # Write-Host "...[Get-VSkRemoteComputerInfo]..."
            $null = Start-Job -Name "$($Computer)" -Scriptblock ${Function:Get-VSkRemoteComputerInfo} -ArgumentList $Computer 
        }

        if ( $ScriptBlockName -eq 'Set-CompWindowsUpdate' ) { 
            $null = Start-Job -Scriptblock {
                param (
                    $Computer,
                    $Update,
                    $AutoReboot,
                    $ImportedFunction
                )

                Invoke-Command -ComputerName $Computer -Scriptblock {
                    param($Update, $AutoReboot, $ImportedFunction)

                    # Write-Host "[Set-CompWindowsUpdate] Update:[$Update], AutoReboot[$AutoReboot]"
                    [ScriptBlock]::Create($ImportedFunction).Invoke($Update, $AutoReboot)

                } -ArgumentList $Update, $AutoReboot, $ImportedFunction
            
            } -ArgumentList ( $Computer, $Update, $AutoReboot, ${Function:Set-CompWindowsUpdate} )
        }
        if ( $ScriptBlockName -eq 'Install' ) { 
            # Install-CompProgram.ps1 -ComputerName EX00001 -InstallPath 'D:\install\7-zip\7z1900-x64.msi'
            $null = Start-Job -Scriptblock ${Function:Install-CompProgram} -ArgumentList $Computer, $InstallPath.FullName
        }
        if ( $ScriptBlockName -eq 'Uninstall' ) { 
            # Uninstall-CompProgram.ps1 -ComputerName EX00001 -CryptedIdNumber 'ASL535LKJAFAFKLKNDG0983095MM36NL3NKLKWEJTBL'
            $null = Start-Job -Scriptblock ${Function:Uninstall-CompProgram}  -ArgumentList $Computer, $UninstallEncryptedParameter
        }
        if ( $ScriptBlockName -eq 'WakeOnLan' ) { 
            $null = Start-Job -Scriptblock ${Function:Invoke-CompWakeOnLan}  -ArgumentList $Computer, $DataArchiveFile
        }
        Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
    }
    While (Get-Job -State "Running") {
        Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
        Start-Sleep 10
    }

    #Get information from each job.
    foreach ( $job in Get-Job ) {
        $Output += @(Receive-Job -Id ($job.Id))
    }
    $null = Stop-Watch -Timer $jobWatch -Name JOBers
    Get-Job | Remove-Job

    if ( $Output.Count -gt 0 ) {
        $Output
    }
    else {
        $False
    }
}
