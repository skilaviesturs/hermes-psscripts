
#region Get-VerifyComputers
Function Get-VerifyComputers {
    Param(
        [Parameter(Position = 0, Mandatory)]
        [string[]]$ComputerNames
    )
    
    $VerifyCompsResultFile = "$($Script.DataDir)\VerifyCompsResult.dat"
        
    Write-Host "[VerifyComp] got for testing [$($ComputerNames.Count)] $(if($ComputerNames.Count -eq 1){"computer"}else{"computers"})"
    Write-Verbose "[Function:VerifyComputers]:got:[$($ComputerNames)]"

    # nosūtam uzdevumu uz remote datoriem

    $JobResults = Set-Jobs -Computers $ComputerNames -ScriptBlockName 'SBVerifyComps'

    # Analizējam saņemtos rezultātus

    $JobResults = $JobResults | Where-Object -Property Computer -ne $null

    if ( $JobResults -ne $False -or
         $JobResults -notlike 'False' -or 
         $JobResults.Count -gt 0 ) {

        $DelComps = @()

        #Analizējam JObbu atgrieztos rezultātus, ievietojam DelComps
        $JobResults | Foreach-Object {
            if ( -not $_.CatchErr ) {
                if ($_.isPingTest -eq $False ) {
                    $DelComps += $_.Computer
                    Write-msg -log -text "[JobResults] computer [$($_.Computer)] is not accessible"
                }#endif
                else {
                    if ($_.isPSversion -eq $False ) { 
                        $DelComps += $_.Computer
                        Write-msg -log -text "[JobResults] [$($_.Computer)] Powershell version is less than 5.0"
                        Write-Verbose "[JobResults] [$($_.Computer)] $($_.msgCatchErr)"
                    }#endif
                    elseif ($_.isPSWUModule -eq $False ) { 
                        $DelComps += $_.Computer
                        Write-msg -log -text "[JobResults] $($_.msgPSWUModuleInf)"
                    }#endif
                    elseif ($_.isJoinObject -eq $False ) {
                        $DelComps += $_.Computer 
                        Write-msg -log -text "[JobResults] $($_.msgJoinObjectInf)"
                    }#endif
                    elseif ($_.isLanguage -eq $False ) {
                        $DelComps += $_.Computer 
                        Write-msg -log -bug -text "[JobResults] $($_.msgLanguageBug)"
                    }#endif
                    elseif ($_.isPolicy -eq $False ) {
                        $DelComps += $_.Computer 
                        Write-msg -log -bug -text "[JobResults] $($_.msgPolicyBug)"
                    }#endif
                }#endelse
            }#endif
            else {
                Write-Host "[JobResults] $($_.CatchErr)"
                Write-msg -log -bug -text "[JobResults] $($_.CatchErr)" 
            }#endelse
        }#endforeach

        #Atbrīvojamies no dublikātiem, ja tādu ir
        $DelComps = $DelComps | Get-Unique

        #Parsējam input masīvu un papildinām DelComps ar dzēšamajiem datoriem, kas nav atbildējuši uz ping
        $ComputerNames | ForEach-Object {
            if ( $JobResults.Computer.Contains($_) -eq $False ) {
                $DelComps += $_
            }#endif
        }#endforeach

        #Aizvācam no input masīva visus datorus, kas nav izturējuši pārbaudi
        $DelComps | ForEach-Object { 
            if ( $ComputerNames.Contains($_) ) {
                $ComputerNames = $ComputerNames -ne $_
                Write-msg -log -text "[VerifyComputers] Computer [$_] is not ready PSRemote. Removed."
            }#endif
        }#endforeach
    }#endif
    else {
        Write-msg -log -bug -text "[JobResults] Oopps!! Jober returned nothing." 
    }#endelse
    if ( $JobResults.GetType().BaseType.name -eq 'Object' ) {
        $tmpJobResults = $JobResults.psobject.copy()
        $JobResults = @()
        $JobResults += @($tmpJobResults)
    }#endif
    if ( $JobResults.count -gt 0) {
        Write-Verbose "[Function:VerifyComputers]:return:[$($ComputerNames)]"
        $JobResults | Export-Clixml -Path $VerifyCompsResultFile -Depth 10 -Force
        return $ComputerNames
    }#endif
    else {
        return $false
    }
}
#endregion

#region Set-Jobs
Function Set-Jobs {
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string[]]$Computers,
        [Parameter(Position = 1, Mandatory = $true)]
        [string]$ScriptBlockName,
        [Parameter(Position = 2, Mandatory = $false)]
        [string]$Argument1
    )
    <# ---------------------------------------------------------------------------------------------------------
    #	Definējam komandu blokus:
    #	SBVerifyComps : pārbaudam datora gatavību darbam ar skriptu
    --------------------------------------------------------------------------------------------------------- #>
    $SBVerifyComps = {
        try {
            $Computer = $args[0]
            #Write-host "[Comp,block1] $Computer"
            If ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop ) {
                $isPingTest = $True
                $RemSess = New-PSSession -ComputerName $Computer -ErrorAction Stop
                if ( ( Invoke-Command -Session $RemSess -ScriptBlock { $PSVersionTable.PSVersion.Major } ) -ge 5 ) { $isPSversion = $True }
        
                #check module 'PSWindowsUpdate' is installed, if not, copy from script's root directory to remote computer
                if (-not ( Invoke-Command -Session $RemSess -ScriptBlock { Get-Module -ListAvailable -Name 'PSWindowsUpdate' } )) {
                    Copy-Item "lib\modules\PSWindowsUpdate\" -Destination "C:\Program Files\WindowsPowerShell\Modules\" -ToSession $RemSess -Recurse -ErrorAction Stop
                    $msgPSWUModuleInf = "PSWindowsUpdate module installed on [$computer]."
                    $isPSWUModule = $True
                }#endif
                else {
                    $isPSWUModule = $True
                }#endelse
                if (-not ( Invoke-Command -Session $RemSess -ScriptBlock { Get-Module -ListAvailable -Name 'Join-Object' } )) {
                    Copy-Item "lib\modules\Join-Object\" -Destination "C:\Program Files\WindowsPowerShell\Modules\" -ToSession $RemSess -Recurse -ErrorAction Stop
                    $msgJoinObjectInf = "Join-Object module installed on [$computer]."
                    $isJoinObject = $True
                }#endif
                else {
                    $isJoinObject = $True
                }#endelse
                #cheking computer is set ExecutionPolicy = RemoteSigned and LanguageMode = FullLanguage
                $psLanguage = Invoke-Command -Session $RemSess -ScriptBlock { $ExecutionContext.SessionState.LanguageMode }
                if ( $psLanguage.value -notlike 'FullLanguage' ) {
                    $msgLanguageBug = "Computer [$Computer] is not ready for PSRemote: LanguageMode is set [$($psLanguage.value)]"
                }#endif
                else {
                    $isLanguage = $True
                }#endelse
                $policy = Invoke-Command -Session $RemSess -ScriptBlock { Get-ExecutionPolicy }
                if ( $policy.value -notlike 'Unrestricted' -and $policy.value -notlike 'RemoteSigned' ) {
                    $msgPolicyBug = "Computer [$Computer] is not ready for PSRemote: policy is set [$($policy.value)]"
                }#endif
                else {
                    $isPolicy = $True
                }#endelse
                if ( $RemSess.count -gt 0 ) {
                    Remove-PSSession -Session $RemSess
                }#endif
            }#endif
            else {
                $msgCatchErr = "Computer [$Computer] is not accessible."
            }#endelse
            $ObjectReturn = New-Object -TypeName psobject -Property @{
                Computer         = $Computer ;
                isPingTest       = if ( $isPingTest ) { $isPingTest } else { $False } ;
                isPSversion      = if ( $isPSversion ) { $isPSversion } else { $False } ;
                isPSWUModule     = if ( $isPSWUModule ) { $isPSWUModule } else { $False } ;
                msgPSWUModuleInf	= if ( $msgPSWUModuleInf ) { $isPSWUModule } else { $null } ;
                isJoinObject     = if ( $isJoinObject ) { $isJoinObject } else { $False } ;
                msgJoinObjectInf	= if ( $msgJoinObjectInf ) { $msgJoinObjectInf } else { $null } ;
                isLanguage       = if ( $isLanguage ) { $isLanguage } else { $False } ;
                msgLanguageBug   = if ( $msgLanguageBug ) { $msgLanguageBug } else { $null } ;
                isPolicy         = if ( $isPolicy ) { $isPolicy } else { $False } ;
                msgPolicyBug     = if ( $msgPolicyBug ) { $msgPolicyBug } else { $null } ;
                msgCatchErr      = if ( $msgCatchErr ) { $msgCatchErr } else { $null } ;
            }#endobject
            return $ObjectReturn
        }#endtry
        catch {
            $msgCatchErr = "$_"
            $ObjectReturn = New-Object -TypeName psobject -Property @{
                Computer         = $Computer ;
                isPingTest       = if ( $isPingTest ) { $isPingTest } else { $False } ;
                isPSversion      = if ( $isPSversion ) { $isPSversion } else { $False } ;
                isPSWUModule     = if ( $isPSWUModule ) { $isPSWUModule } else { $False } ;
                msgPSWUModuleInf	= if ( $msgPSWUModuleInf ) { $isPSWUModule } else { $null } ;
                isJoinObject     = if ( $isJoinObject ) { $isJoinObject } else { $False } ;
                msgJoinObjectInf	= if ( $msgJoinObjectInf ) { $msgJoinObjectInf } else { $null } ;
                isLanguage       = if ( $isLanguage ) { $isLanguage } else { $False } ;
                msgLanguageBug   = if ( $msgLanguageBug ) { $msgLanguageBug } else { $null } ;
                isPolicy         = if ( $isPolicy ) { $isPolicy } else { $False } ;
                msgPolicyBug     = if ( $msgPolicyBug ) { $msgPolicyBug } else { $null } ;
                msgCatchErr      = if ( $msgCatchErr ) { $msgCatchErr } else { $null } ;
            }#endobject
            if ( $RemSess.count -gt 0 ) {
                Remove-PSSession -Session $RemSess
            }#endif
            return $ObjectReturn
        }#endcatch
    }#endblock

    <# ---------------------------------------------------------------------------------------------------------
    # SBWindowsUpdate: izsaucam Windows update uz attālinātās darbstacijas
    ---------------------------------------------------------------------------------------------------------#>
    $SBWindowsUpdate = {
        $Computer = $args[0]
        $WinUpdFile = $args[1]
        $Update = $args[2]
        $AutoReboot = $args[3]
        $OutputResults = Invoke-Command -ComputerName $Computer -FilePath $WinUpdFile -ArgumentList ($Update, $AutoReboot)
        return $OutputResults
    }#endblock

    <# ---------------------------------------------------------------------------------------------------------
    # SBInstall: izsaucam programmas uzstādīšanas procesu
    # Set-CompProgram.ps1 [-ComputerName] <string> [-InstallPath <FileInfo>] [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    $SBInstall = {
        $Computer = $args[0]
        $CompProgramFile = $args[1]
        $Install = $args[2]
        $OutputResults = Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-InstallPath $Install "
        return $OutputResults
    }#endblock

    <# ---------------------------------------------------------------------------------------------------------
    # SBInstall: izsaucam programmas noņemšanas procesu
    # Set-CompProgram.ps1 [-ComputerName] <string> [-CryptedIdNumber <string>] [<CommonParameters>]
    ---------------------------------------------------------------------------------------------------------#>
    $SBUninstall = {
        $Computer = $args[0]
        $CompProgramFile = $args[1]
        $EncryptedParameter = $args[2]
        $OutputResults = Invoke-Expression "& `"$CompProgramFile`" `-ComputerName $Computer `-CryptedIdNumber $EncryptedParameter "
        return $OutputResults
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
        $OutputResults = Invoke-Expression "& `"$CompWakeOnLanFile`" `-ComputerName $Computer `-DataArchiveFile `"$DataArchiveFile`" `-CompTestOnline `"$CompTestOnlineFile`" "
        return $OutputResults
    }#endblock
    <# ---------------------------------------------------------------------------------------------------------
        [JOBers] kods
    --------------------------------------------------------------------------------------------------------- #>
    $jobWatch = [System.Diagnostics.Stopwatch]::startNew()
    $Output = @()
    Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
    ForEach ( $Computer in $Computers ) {
        While ($(Get-Job -state running).count -ge $Script:MaxJobsThreads) {
            Start-Sleep -Milliseconds 10
        }#endWhile
        if ( $ScriptBlockName -eq 'SBVerifyComps' ) { 
            $null = Start-Job -Name "$($Computer)" -Scriptblock $SBVerifyComps -ArgumentList $Computer 
        }#endif
        if ( $ScriptBlockName -eq 'SBWindowsUpdate' ) { 
            $null = Start-Job -Name "$($Computer)" -Scriptblock $SBWindowsUpdate -ArgumentList $Computer, $WinUpdFile, $Update, $AutoReboot 
        }#endif
        if ( $ScriptBlockName -eq 'SBInstall' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile, $Install"
            $null = Start-Job -Scriptblock $SBInstall -ArgumentList $Computer, $CompProgramFile, $Install
        }#endif
        if ( $ScriptBlockName -eq 'SBUninstall' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1"
            $null = Start-Job -Scriptblock $SBUninstall -ArgumentList $Computer, $CompProgramFile, $Argument1
        }#endif
        if ( $ScriptBlockName -eq 'SBWakeOnLan' ) { 
            Write-Verbose "[StartJob] Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile"
            $null = Start-Job -Scriptblock $SBWakeOnLan -ArgumentList $Computer, $CompWakeOnLanFile, $DataArchiveFile, $CompTestOnlineFile
        }#endif
        Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
    }#endForEach
    While (Get-Job -State "Running") {
        Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
        Start-Sleep 10
    }
    #Get information from each job.
    foreach ( $job in Get-Job ) {
        $result = @()
        $result = Receive-Job -Id ($job.Id)
        if ( $result -or $result.count -gt 0 ) {
            $Output += $result
        }#endif
    }#endforeach
    Stop-Watch -Timer $jobWatch -Name JOBers
    Get-Job | Remove-Job

    if ( $Output.Count -gt 0 ) {
        Return $Output
    }#endif
    else {
        Return $False
    }#endelse
}#endOfFunction
#endregion

#region Get-NormaliseDiskLabelsForExcel
Function Get-NormaliseDiskLabelsForExcel {
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Computers
    )
    $TemplateDisks = [PSCustomObject][ordered]@{}
    #Get uniqe disk label's collection from computers and write in array
    foreach ($computer in $computers) {
        foreach ( $Item in $computer.PsObject.Properties ) {
            if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize' ) {
                if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
                    $TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
                }
            }
            if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize\sFree' ) {
                if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
                    $TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
                }
            }
        }
    }
    #Add to each computer's properties missing disk label
    foreach ( $computer in $computers ) {
        foreach ($label in $TemplateDisks.PsObject.Properties) {
            if ( -not ( Get-Member -InputObject $computer -Name $label.Name ) ) {
                Add-Member -InputObject $computer -NotePropertyName $label.Name -NotePropertyValue 'none' -Force
            }
        }
    }
    return $computers
}
#endregion