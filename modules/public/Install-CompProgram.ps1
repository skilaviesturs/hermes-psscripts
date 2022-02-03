#Requires -Version 5.1
Function Install-CompProgram {
    <#
    .SYNOPSIS
    Attālināta programmatūras uzstādīšana un noņemšana
    
    .DESCRIPTION
    Skripts nodrošina uz attālinātā datora
    msi vai exe pakotnes uzkopēšanu un uzstādīšanu
    
    .PARAMETER ComputerName
    Norādam attālinātā datora NETBIOS vai DNS vārdu
    
    .PARAMETER InstallPath
    Norādam uzstādāmās programmatūras MSI vai EXE pakotnes atrašanās vietu.
    Lietotājam, ar kuru veicam skripta darbināšanu, jābūt pilnām tiesībām uz pakotni.
    
    .PARAMETER Help
    Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.
    
    .EXAMPLE
    Install-CompProgram.ps1 -ComputerName EX00001 -InstallPath 'D:\install\7-zip\7z1900-x64.msi'
    Uzstādam uz datora EX00001 programmatūras instalācijas pakotni 7z1900-x64.msi
    
    .NOTES
        Author:	Viesturs Skila
        Version: 1.2.3
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory,
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
    
        [Parameter(Position = 1, Mandatory,
            HelpMessage = "Path of installer msi or exe file.")]
        [ValidateScript( {
                if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
                    Write-Host "File does not exist"
                    throw
                }
                if ( $_ -notmatch ".msi|.exe") {
                    Write-Host "`nThe file specified in the path argument must be msi file`n" -ForegroundColor Yellow
                    throw
                }
                return $True
            } ) ]
        [System.IO.FileInfo]$InstallPath,
    
        [Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
        [switch]$Help = $False
    )
    BEGIN {
        if (-not $PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
        }
        Write-Host "[Install-CompProgram] ------------------------------------------------------------"
        Write-Host "[Install-CompProgram]:ComputerName[$ComputerName]"
        Write-Host "[Install-CompProgram]:InstallPath [$InstallPath]"

        <# ---------------------------------------------------------------------------------------------------------
        Skritpa konfigurācijas datnes
        --------------------------------------------------------------------------------------------------------- #>
        $CurVersion = "1.2.3"
        #$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
        $__ScriptName = $MyInvocation.MyCommand
        $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
        #$LogFileDir = "log"
        # $LogMessage = @()
        $LogMessage = [System.Collections.ArrayList]@()
        $WorkingDirectory = 'C:\temp'

        if ($Help) {
            Write-Host "`nVersion:[$CurVersion]`n"
            $text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
            $text | ForEach-Object { Write-Host $($_) }
            Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
            Exit
        }
        
        # Pārbaudam uz mērķa datora brīvo vietu un izveidojam pagaidu instalācijas mapi, kurā iekopēsim msi
        $CheckSpace = {
            param(
                [Parameter(Position = 0)]
                [string]$tempPath,
                [Parameter(Position = 1)]
                [int]$packageSize
            )
    
            #region atrodam c: diska apjomu un brīvo vietu
    
            $FreeSpace = 0
            $devices = Get-CimInstance -ClassName win32_LogicalDisk -Filter "DriveType = '3'" -Property DeviceID, Size, FreeSpace
            foreach ( $disk in $devices ) {
                if ( $disk.DeviceID -like "C*" ) {
                    $FreeSpace = $disk.FreeSpace
                }
            }
    
            #endregion
    
            #region ja vietas pietiekoši, izveidojam pagaidu mapi un atgriežam rezultātu
    
            if ( $FreeSpace -gt ( $packageSize * 2 ) -or $FreeSpace -gt 2147483648 ) {
                if ( -NOT ( Test-Path -Path $tempPath ) ) {
                    $null = New-Item -Path $tempPath -ItemType 'Directory' -Force
                }
                return "Ok"
            }
            else {
                return "There's no enough free space [$($FreeSpace/1GB)]GB. Need at least [$( if( ($packageSize * 2) -gt 2147483648 ) { ($packageSize / 1GB) * 2 } else {"2"})]GB free space."
            }
    
            #endregion
        }
    }

    PROCESS {

        try {
            $RemoteSession = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
        
            if ( $RemoteSession.Count -gt 0 ) {

                try {

                    $InstallFile = Get-ChildItem -Path $InstallPath -ErrorAction Stop

                    Write-Host "[Install-CompProgram]:InstallFile.Length [$($InstallFile.Length)]"

                    
                    # padodam: lokālā diska tmp mapi, msi pakotnes izmēru
                    $parametersCheck = @{
                        Session      = $RemoteSession
                        ScriptBlock  = $CheckSpace
                        ArgumentList = $WorkingDirectory, $InstallFile.Length
                        ErrorAction  = 'Stop'
                    }
                    try {
                    
                        $result = Invoke-Command @parametersCheck
                    }
                    catch {
                        $LogMessage += $_ | Out-String | ForEach-Object { @( "[CheckDiskSize] $_" ) }
                    }

                    Write-Host "[Install-CompProgram]:result:[$result]"
                    if ( $result -eq 'Ok' ) {
                        #region Kopējam msi pakotni uz remote datora mapi
                        $parametersCopy = @{
                            Path        = $InstallFile.FullName
                            Destination = $WorkingDirectory
                            ToSession   = $RemoteSession
                            Force       = $true
                            ErrorAction = 'Stop'
                        }
                        try {
                            
                            Copy-Item @parametersCopy
                        }
                        catch {
                            $LogMessage += $_ | Out-String | ForEach-Object { @( "[CopyTo] $_" ) }
                        }
                        #endregion

                        Write-Host "[Install-CompProgram]:InstallFile.Name:[$($InstallFile.Name)], WorkingDirectory:[$WorkingDirectory]"
                        # padodam: lokālā diska tmp mapi, msi pakotnes datnes nosaukumu
                        $parametersInstall = @{
                            Session      = $RemoteSession
                            # ScriptBlock  = ${Function:Invoke-VSkInstallBlock}
                            # ArgumentList = "$WorkingDirectory", "$($InstallFile.Name)"
                            ErrorAction  = 'Stop'
                        }
                        try {

                            $InstallResult = Invoke-Command @parametersInstall -ScriptBlock {
                                param($Path, $FileName, $ImportedFunction)
                                
                                [ScriptBlock]::Create($ImportedFunction).Invoke($Path, $FileName)

                            } -ArgumentList $WorkingDirectory, $InstallFile.Name, ${Function:Invoke-VSkInstallBlock}

                            $InstallResult |  ForEach-Object { $LogMessage += @( "$_") }
                        }
                        catch {
                            $LogMessage += $_ | Out-String | ForEach-Object { @( "[VSkInstallBlock] $_" ) }
                        }
                    }
                    else {
                        $LogMessage += @($result)
                    }
                }
                catch {
                    $LogMessage += $_ | Out-String | ForEach-Object { @( "2[Install-CompProgram] $_" ) }
                }
            }

        }
        catch {
            $LogMessage += $_ | Out-String | ForEach-Object { @( "3[Install-CompProgram] $_" ) }
        }
        finally {
            if ( $RemoteSession.Count -gt 0 ) {
                Remove-PSSession -Session $RemoteSession
            }
        }

    }

    END {

        $Output = @()
        $i = 0
        $LogMessage | ForEach-Object {
            $Output += @( New-Object -TypeName psobject -Property @{
                    id       = $i
                    Computer	= [string]$ComputerName;
                    Message  = [string]$_;
                }
            )
            $i++
        }

        $Output
    }
}