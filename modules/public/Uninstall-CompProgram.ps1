#Requires -Version 5.1
Function Uninstall-CompProgram {
    <#
    .SYNOPSIS
    Attālināta programmatūras uzstādīšana un noņemšana
    
    .DESCRIPTION
    Skripts nodrošina uz attālinātā datora:
    programmas noņemšanu
    
    .PARAMETER ComputerName
    Norādam attālinātā datora NETBIOS vai DNS vārdu
    
    .PARAMETER UninstallIdNumber
    Norādam uzstādāmās programmatūras unikālo identifikatoru.
    Programmatūras identifikatoru varam iegūt ar skriptu Get-CompSoftware.ps1
    
    .PARAMETER Help
    Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.
    
    .EXAMPLE
    Uninstall-CompProgram.ps1 EX00001 -UninstallIdNumber '{23170F69-40C1-2702-1900-000001000000}'
    Noņemam programmatūras instalāciju, kuras identifikācijas numurs ir {23170F69-40C1-2702-1900-000001000000}
    
    .EXAMPLE
    Uninstall-CompProgram.ps1 EX00001 -CryptedIdNumber 'ASL535LKJAFAFKLKNDG0983095MM36NL3NKLKWEJTBL'
    Noņemam programmatūras instalāciju, kuras identifikācijas numurs ir kriptēts, lai tas saturētu {}.
    Šifrēto parametru skripts atšifrēs.

    .NOTES
        Author:	Viesturs Skila
        Version: 1.2.3
    #>
    [CmdletBinding(DefaultParameterSetName = 'UninstallCrypt')]
    Param(
        [Parameter(Position = 0, Mandatory = $True,
            ParameterSetName = 'Uninstall',
            HelpMessage = "Name of computer")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                if ( [String]::IsNullOrWhiteSpace($_) ) {
                    Write-Host "`nEnter the name of computer`n" -ForegroundColor Yellow
                    throw
                }
                return $True
            } ) ]
        [Parameter(Position = 0, Mandatory = $True,
            ParameterSetName = 'UninstallCrypt',
            HelpMessage = "Name of computer")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                if ( [String]::IsNullOrWhiteSpace($_) ) {
                    Write-Host "`nEnter the name of computer`n" -ForegroundColor Yellow
                    throw
                }
                return $True
            } ) ]
        [string]$ComputerName,
        
        # {23170F69-40C1-2702-1900-000001000000}
        [Parameter(Position = 1, Mandatory = $True,
            ParameterSetName = 'Uninstall',
            HelpMessage = "Identifying number of program you want to uninstall")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                if ( [String]::IsNullOrWhiteSpace($_) ) {
                    Write-Host "`nEnter Identifying Number of program`n" -ForegroundColor Yellow
                    throw
                }
                return $True
            } ) ]
        [string]$UninstallIdNumber,
    
        # Crypted IdentifyingNumber
        [Parameter(Position = 1, Mandatory = $True,
            ParameterSetName = 'UninstallCrypt',
            HelpMessage = "Crypted Identifying number of program you want to uninstall")]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                if ( [String]::IsNullOrWhiteSpace($_) ) {
                    Write-Host "`nEnter Crypted Identifying Number of program`n" -ForegroundColor Yellow
                    throw
                }
                return $True
            } ) ]
        [string]$CryptedIdNumber,
    
        [Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
        [switch]$Help = $False
    )
    BEGIN {
        <# ---------------------------------------------------------------------------------------------------------
        Skritpa konfigurācijas datnes
        --------------------------------------------------------------------------------------------------------- #>
        $CurVersion = "1.2.3"
        #$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()

        $LogMessage = @()
        
        # atinstalējam programmatūru uz mērķa datoru
        $Uninstall = {
            [CmdletBinding()]
            param(
                [Parameter(Position = 0)]
                [string]$Number
            )
    
            $LogMessage = @()
            
            $programm = Get-WmiObject -ClassName 'Win32_Product' | Where-Object { $_.IdentifyingNumber -eq $Number }
            
            if ($programm) {
                $null = $programm.Uninstall()
                $LogMessage += @("[Uninstaller] [SUCCESS] by default uninstaller.")
            }
            else {
                $LogMessage += @("[Uninstaller] [WARN] Win32_Product did not find anything.")
                
                #region nevarējām atrast programmu ar WmiOjectu, meklējam reģistros ierakstus pēc IdentifyNumber
    
                $RegistryPaths = @(
                    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                    'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
                )
    
                $SchTaskCommand = $null
    
                foreach ( $path in $RegistryPaths ) {
                    $registryObject = Get-ItemProperty -Path "$path\$Number" -ErrorAction SilentlyContinue
                    if ( $registryObject ) {
                        $QuietSchTaskCommand = $registryObject.QuietUninstallString
                        $SchTaskCommand = $registryObject.UninstallString
                        $WorkingDirectory = $registryObject.InstallLocation
                        $LogMessage += @("[Uninstaller] [INFO] found QuietUninstallString[$QuietSchTaskCommand]:UninstallString[$SchTaskCommand] in [$path\$Number]")
                        break
                    }
                }
    
                #endregion
                
                if ( $QuietSchTaskCommand -or $SchTaskCommand ) {
                    #$SchTaskCommand = $SchTaskCommand.Replace('"', "'")
                    #Write-Host "[Uninstaller]:found QuietUninstallString[$SchTaskCommand]"
                    
                    try {
                        
                        #region ievietojam tekošo lietotāju servera ForJobs grupā, lai lietotājs var izpildīt fonā
                        <#
                            try {
                                $ForJobs = [ADSI]"WinNT://$env:ComputerName/ForJobs,group"
                                $User = [ADSI]"WinNT://$((whoami).replace("\","/")),user"
                                $ForJobs.Add($User.Path)
                            }#endtry
                            catch {
                                Write-Warning $($_.Exception.Message)
                            }
                            #>
                        #endregion
                        
                        #region formējam Uninstall ScheduledTask
                        
                        $__RandomID	= Get-Random -Minimum 100000 -Maximum 999999
                        $taskName = "[Uninstaller] [$Number] [$__RandomID]"
                        $logFile = "C:/temp/uninstaller-$Number-$__RandomID.log"
                        
                        if ( $QuietSchTaskCommand ) {
                            $Argument = "`/C $QuietSchTaskCommand /L*V $logFile"
                        }
                        elseif ( $SchTaskCommand ) {
                            $Argument = "`/C `"$SchTaskCommand`" /S /L*V $logFile"
                        }
    
                        $LogMessage += @("[Uninstaller] [INFO] TaskName:[$taskName]; Argument: [$Argument]")

                        $action = New-ScheduledTaskAction `
                            -Execute 'cmd.exe' `
                            -Argument $Argument `
                            -WorkingDirectory "$WorkingDirectory\" `
                            -ErrorAction Stop
                        
                        #izveidojam trigeri
                        $trigger = New-ScheduledTaskTrigger `
                            -Once `
                            -At ([DateTime]::Now.AddMinutes(1)) `
                            -ErrorAction Stop
    
                        #izveidojam lietotāju
                        # Accepted values: "BUILTIN\Administrators", "SYSTEM", "$(whoami)""
                        # Accepted values LogonType: None, Password, S4U, Interactive, Group, ServiceAccount, InteractiveOrPassword
                        $principal = New-ScheduledTaskPrincipal `
                            -UserID "SYSTEM" `
                            -LogonType S4U `
                            -RunLevel Highest `
                            -ErrorAction Stop
                        #-GroupID 'BUILTIN\Administrators' `
    
                        #liekam kopā un izveidojam task objektu
                        $null = Register-ScheduledTask `
                            -TaskName $taskName `
                            -Action $action `
                            -Trigger $trigger `
                            -Principal $principal `
                            -Description "Automated task set by script" `
                            -ErrorAction Stop
    
                        #Papildinām task objekta parametrus
                        $TargetTask = Get-ScheduledTask -ErrorAction Stop |
                        Where-Object -Property TaskName -eq $taskName
    
                        $TargetTask.Author = $__ScriptName
                        $TargetTask.Triggers[0].StartBoundary = [DateTime]::Now.AddMinutes(1).ToString("yyyy-MM-dd'T'HH:mm:ss")
                        $TargetTask.Triggers[0].EndBoundary = [DateTime]::Now.AddHours(1).ToString("yyyy-MM-dd'T'HH:mm:ss")
                        $TargetTask.Settings.AllowHardTerminate = $True
                        $TargetTask.Settings.DeleteExpiredTaskAfter = 'PT0S'
                        $TargetTask.Settings.ExecutionTimeLimit = 'PT1H'
                        #Accepted values: Parallel, Queue, IgnoreNew
                        $TargetTask.Settings.MultipleInstances = 'IgnoreNew'
                        $TargetTask.Settings.volatile = $False
    
                        #Papildināto objektu saglabājam
                        $TargetTask | Set-ScheduledTask -ErrorAction Stop | Out-Null
                        $LogMessage += @("[Uninstaller] [SUCCESS] sheduled task [$TaskName]: successfully created ")

                        #endregion
                    }
                    catch {
                        $LogMessage += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }
                }
                else {
                    $LogMessage += @("[Uninstaller] [WARN] did not find anything in [Registry].")
                }
            }
            
            $LogMessage
        }
    }
    
    PROCESS {
    
        try {
            $RemoteSession = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
            
            if ( $RemoteSession.Count -gt 0 ) {
                
                if ( $PSCmdlet.ParameterSetName -like "Uninstall*" ) {
                    try {
    
                        #region atšifrējam parametru

                        if ( $PSCmdlet.ParameterSetName -eq "UninstallCrypt" ) {
    
                            $secParameter = $CryptedIdNumber | ConvertTo-SecureString
                            $Marshal = [System.Runtime.InteropServices.Marshal]
                            $Bstr = $Marshal::SecureStringToBSTR($secParameter)
                            $UninstallIdNumber = $Marshal::PtrToStringAuto($Bstr)
                            $Marshal::ZeroFreeBSTR($Bstr)
                            $LogMessage += @("[Uninstaller] [INFO] decrypted IDNumber:[$UninstallIdNumber]")
                        }

                        #endregion
                            
                        #region padodam: programmas identifikācijas numuru

                        $parameters = @{
                            Session      = $RemoteSession
                            ScriptBlock  = $Uninstall
                            ArgumentList = ( $UninstallIdNumber )
                            ErrorAction  = 'Stop'
                        }
                        $UninstallResult = Invoke-Command @parameters
                        $UninstallResult | ForEach-Object { $LogMessage += @($_) }
    
                        #endregion
                    }
                    catch {
                        $LogMessage += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }
                }
            }
            else {}
        }
        catch {
            $LogMessage += $_ | Out-String | ForEach-Object { @( "$_" ) }
        }
        finally {
            if ( $RemoteSession.Count -gt 0 ) {
                Remove-PSSession -Session $RemoteSession
            }
        }
    }
    
    END {
    
        #region Make structured message object

        $Output = @()
        $i = 0
        $LogMessage | ForEach-Object {
            $Output += @( New-Object -TypeName psobject -Property @{
                    id       = $i
                    Computer = [string]$ComputerName
                    Message  = [string]$_
                }
            )
            $i++
        }

        #endregion
        $Output
    }
}