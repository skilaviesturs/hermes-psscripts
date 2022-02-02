#Requires -Version 5.1
Function Send-VSkJob {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory)]
        [string[]]$Computers,
        [Parameter(Position = 1, Mandatory)]
        [string]$ScriptBlockName,
        [switch]$Update,
        [switch]$AutoReboot
    )
    $MaxJobsThreads = 25

    <# ---------------------------------------------------------------------------------------------------------
    # SBInstall: izsaucam programmas uzstādīšanas procesu
    # Set-CompProgram.ps1 [-ComputerName] <string> [-InstallPath <FileInfo>] [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    $SBInstall = {
        $Computer = $args[0]
        $CompProgramFile = $args[1]
        $Install = $args[2]
        Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-InstallPath $Install "
    }#endblock

    <# ---------------------------------------------------------------------------------------------------------
    # SBInstall: izsaucam programmas noņemšanas procesu
    # Set-CompProgram.ps1 [-ComputerName] <string> [-CryptedIdNumber <string>] [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    $SBUninstall = {
        $Computer = $args[0]
        $CompProgramFile = $args[1]
        $EncryptedParameter = $args[2]
        Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-CryptedIdNumber $EncryptedParameter "
    }#endblock

    <# ---------------------------------------------------------------------------------------------------------
    # SBWakeOnLan: izsaucam programmas noņemšanas procesu
    # Invoke-CompWakeOnLan.ps1 [-ComputerName] <string[]> [-DataArchiveFile] <FileInfo> [-CompTestOnline] <FileInfo> [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    $SBWakeOnLan = {
        $Computer = $args[0]
        $CompWakeOnLanFile = $args[1]
        $DataArchiveFile = $args[2]
        $CompTestOnlineFile = $args[3]
        Invoke-Expression "& `"$CompWakeOnLanFile`" `-ComputerName $Computer `-DataArchiveFile `"$DataArchiveFile`" `-CompTestOnline `"$CompTestOnlineFile`" "
    }#endblock
    <# ---------------------------------------------------------------------------------------------------------
    # SBWakeOnLan: izsaucam programmas noņemšanas procesu
    # Invoke-CompWakeOnLan.ps1 [-ComputerName] <string[]> [-DataArchiveFile] <FileInfo> [-CompTestOnline] <FileInfo> [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    # $SetCompWindowsUpdate = {
    #     $Computer = $args[0]
    #     $Update = $args[1]
    #     $AutoReboot = $args[2]
    #     $LocalFunction = $args[3]
    #     Invoke-Command -ComputerName $Computer -Scriptblock {
    #         param($update, $autoReboot, $localFunction)
    #         [ScriptBlock]::Create($localFunction).Invoke($update, $autoReboot)
    #     } -ArgumentList $Update, $AutoReboot, $LocalFunction
    # }
    <# ---------------------------------------------------------------------------------------------------------
        [JOBers] kods
    --------------------------------------------------------------------------------------------------------- #>
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
        if ( $ScriptBlockName -eq 'SBInstall' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile"
            # $null = Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile, $Install
            $null = Start-Job -Scriptblock ${Function:Install-CompProgram} -ArgumentList $Computer, $CompProgramFile
        }
        if ( $ScriptBlockName -eq 'SBUninstall' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1"
            $null = Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1
        }
        if ( $ScriptBlockName -eq 'SBWakeOnLan' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile"
            $null = Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile
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
