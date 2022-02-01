#Requires -Version 5.1
Function Send-VSkJob {
    Param(
        [Parameter(Position = 0, Mandatory)]
        [string[]]$Computers,
        [Parameter(Position = 1, Mandatory)]
        [string]$ScriptBlockName,
        [Parameter(Position = 2)]
        [string]$Argument1
    )
    $MaxJobsThreads = 30
    <# ---------------------------------------------------------------------------------------------------------
    # SBWindowsUpdate: izsaucam Windows update uz attālinātās darbstacijas
    ---------------------------------------------------------------------------------------------------------#>
    # $SBWindowsUpdate = {
    #     $Computer = $args[0]
    #     $WinUpdFile = $args[1]
    #     $Update = $args[2]
    #     $AutoReboot = $args[3]
    #     Invoke-Command -ComputerName $Computer -FilePath $WinUpdFile -ArgumentList ($Update, $AutoReboot)
    # }#endblock

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
        [JOBers] kods
    --------------------------------------------------------------------------------------------------------- #>
    $jobWatch = [System.Diagnostics.Stopwatch]::startNew()
    $Output = @()
    Write-Host -NoNewLine "[Set-VSkJobs] Running jobs : " -ForegroundColor Yellow -BackgroundColor Black

    ForEach ( $Computer in $Computers ) {
        While ($(Get-Job -state running).count -ge $MaxJobsThreads) {
            Start-Sleep -Milliseconds 10
        }
        
        if ( $ScriptBlockName -eq 'SBVerifyComps' ) { 
            Write-Host "[JOBers] Invoke-VSkRemoteComputerInfo..."
            $null = Start-Job -Name "$($Computer)" -Scriptblock ${Function:Invoke-VSkRemoteComputerInfo} -ArgumentList $Computer 
            
        }
        if ( $ScriptBlockName -eq 'SBWindowsUpdate' ) { 
            # $null = Start-Job -Name "$($Computer)" -Scriptblock $SBWindowsUpdate -ArgumentList $Computer, $WinUpdFile, $Update, $AutoReboot 
            $null = Start-Job -Name "$($Computer)" -Scriptblock ${Function:Set-CompUpdate} -ArgumentList $Update, $AutoReboot 
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
        $Output = @(
            Receive-Job -Id ($job.Id)
        )
    }
    Stop-Watch -Timer $jobWatch -Name JOBers
    Get-Job | Remove-Job

    if ( $Output.Count -gt 0 ) {
        Return $Output
    }
    else {
        Return $False
    }
}
