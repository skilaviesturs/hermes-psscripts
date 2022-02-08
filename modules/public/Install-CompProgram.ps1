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
        Version: 1.3.0
    #>
    [CmdletBinding(DefaultParameterSetName = 'Install')]
    Param(
        [Parameter(Position = 0, Mandatory,
            ParameterSetName = 'Install',
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
            ParameterSetName = 'Install',
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
        [switch]$Help
    )
    BEGIN {
        if (-not $PSBoundParameters.ContainsKey('Verbose')) {
            $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
        }
        # Write-Host "[Install-CompProgram] ------------------------------------------------------------"
        # Write-Host "[Install-CompProgram]:ComputerName[$ComputerName]"
        # Write-Host "[Install-CompProgram]:InstallPath [$InstallPath]"
        # Write-Host "[Get-Location]:[$(Get-Location)]"

        <# ---------------------------------------------------------------------------------------------------------
        Skritpa konfigurācijas datnes
        --------------------------------------------------------------------------------------------------------- #>

        $LogMessage = [System.Collections.ArrayList]@()
        $WorkingDirectory = 'C:\temp'
        
        # region Pārbaudam uz mērķa datora brīvo vietu un izveidojam pagaidu instalācijas mapi, kurā iekopēsim msi
        $CheckRemoteSpace = {
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

        #endregion

        $InstallBlock = {
            param(
                [Parameter(Position = 0)]
                [string]$WorkingDirectory,
                [Parameter(Position = 1)]
                [string]$FileName
             )
        
            # Write-Host "[InstallBlock] [$WorkingDirectory]\[$FileName]" 
            $LogMessage = [System.Collections.ArrayList]@()
            # $LogMessage = @()
            $LogMessage += @("[InstallBlock] [INFO] got $WorkingDirectory\$FileName")
        
            #region sesijas lietotājam iestatam kontroli pār pagaidu direktoriju
            #$WorkingDirectory = "C:\temp"
        
            $packagePath = Get-ChildItem -Path "$WorkingDirectory" -Recurse
            $Acl = Get-Acl -Path "$WorkingDirectory"
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$(whoami)", "FullControl", "Allow")
            $Acl.SetAccessRule($AccessRule)
            $packagePath | ForEach-Object { Set-Acl -Path $_.FullName -AclObject $Acl }
        
            #endregion
        
            if ( Test-Path -Path "$WorkingDirectory\$FileName" -Type Leaf  -ErrorAction Stop ) {
        
        
                $Source = Get-ChildItem -Path "$WorkingDirectory\$FileName"
 
                $DataStamp = get-date -Format yyyyMMddTHHmmss
                $logFile = '{0}-{1}.log' -f "$($Source.DirectoryName)\$($Source.BaseName)", $DataStamp
        
                if ( $Source.Extension -eq '.msi' ) {
        
                    $MSIArguments = @(
                        "/i"
                        ('"{0}"' -f $Source.FullName)
                        "/qn"
                        "/norestart"
                        "AGREETOLICENSE=yes"
                        "/L*v"
                        $logFile
                    )
                    $object = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru
                    $LogMessage += @("[InstallBlock] [INFO] [$($Source.FullName)] log file [$logFile]")
        
                    #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                    $patterns = @(
                        '-- Installation completed successfully.'
                        'Reconfiguration success or error status: 0.'
                        'Installation success or error status: 0.'
                    )
                    $patterns | ForEach-Object {
                        if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                            $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                            $LogMessage += @("[InstallBlock] [SUCCESS] $output")
                        }
                    }
                    $patterns = @(
                        'Windows Installer requires a system restart.'
                    )
                    $patterns | ForEach-Object {
                        if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                            $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                            $LogMessage += @("[InstallBlock] [WARN] $output")
                        }
                    }
                }
                elseif ( $Source.Extension -eq '.exe'  ) {
        
                    # $logFile = "$(Split-Path -Path "$($Source.FullName)" -Parent)\$($SourceFile.BaseName).log"
                    $LogMessage += @("[InstallBlock] [INFO] [$($Source.FullName)] log file [$logFile]")
        
                    if ( $Source.BaseName -like "AcroRdrDC*"  ) {
                        $Arguments = "`/c $($Source.FullName) `/sAll `/rs `/msi EULA_ACCEPT=YES `/L`*V $logFile"
                    }
                    else {
                        $Arguments = "`/c $($Source.FullName) `/S `/L`*V $logFile"
                    }

                    # Write-Host "[InstallBlock] Source.FullName:[$($Source.FullName)] log file [$logFile]"
                    # Write-Host "[InstallBlock] Log file:[$logFile]"
                    # # Write-Host "[InstallBlock] [INFO] Source.DirectoryName:[$($(Split-Path -Path "$($Source.FullName)" -Parent))]"
                    # Write-Host "[InstallBlock] Source.DirectoryName:[$($Source.DirectoryName)]"
                    # Write-Host "[InstallBlock] Arguments:[$Arguments]"
                    
                    $InstallObject = New-object System.Diagnostics.ProcessStartInfo -Property @{
                        CreateNoWindow         = $true
                        UseShellExecute        = $false
                        RedirectStandardOutput = $true
                        RedirectStandardError  = $true
                        FileName               = 'cmd.exe'
                        Arguments              = $Arguments
                        WorkingDirectory       = $Source.DirectoryName
                        # WorkingDirectory       = "$(Split-Path -Path "$($Source.FullName)" -Parent)"
                    }
                    $process = New-Object System.Diagnostics.Process 
                    $process.StartInfo = $InstallObject 
                    $null = $process.Start()
                    $output = $process.StandardOutput.ReadToEnd()
                    $outputErr = $process.StandardError.ReadToEnd()
                    $process.WaitForExit() 
                    if ( $output ) { $output | Out-File $logFile -Append }
                    if ( $outputErr ) { $outputErr | Out-File $logFile -Append }
                    
                    # Write-Host "[InstallBlock] ExitCode:[$($process.ExitCode)]"

                    if ($process.ExitCode -eq 0) { 
                        $LogMessage += @("[InstallBlock] [SUCCESS] process successfull")
                    }
                    else { 
                        $LogMessage += @("[InstallBlock] [ERROR] process failed with error code [$($process.ExitCode)]")
                    }
        
                    #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                    $patterns = @(
                        '-- Installation completed successfully.'
                        'Reconfiguration success or error status: 0.'
                        'Installation success or error status: 0.'
                    )

                    if ( Test-Path -Path "$logFile" -Type Leaf  -ErrorAction Stop ) {

                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[InstallBlock] [SUCCESS] $output")
                            }
                        }
                        $patterns = @(
                            'Windows Installer requires a system restart.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[InstallBlock] [WARN] $output")
                            }
                        }
                    }
                }
                else {
                    Write-host "[1]: $_ "
                    $LogMessage += @("[InstallBlock] [ERROR] supports only msi or exe format")
                }
            }
            else {
                Write-host "[2]: $_ "
                $LogMessage += @("[InstallBlock] [ERROR] not found [$WorkingDirectory\$FileName]")
            }
        
            # Write-Host "[InstallBlock]:Source.FullName:[$($Source.FullName)]"
            if ( Test-Path -Path "$($Source.FullName)" -PathType Leaf ) {
                Remove-Item -Path "$($Source.FullName)" -Force
            }
            
            $LogMessage
        }
    }

    PROCESS {

        try {
            $RemoteSession = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
        
            if ( $RemoteSession.Count -gt 0 ) {

                try {

                    $InstallFile = Get-ChildItem -Path $InstallPath -ErrorAction Stop

                    Write-Verbose "[Install-CompProgram]:InstallFile.Length [$($InstallFile.Length)]"
                    # Write-Host "[Install-CompProgram] Invoke-VSkInstallBlock aviable [$(if ( ( Get-Command -Module VSKWinUpdate ).name -like 'Invoke-VSkInstallBlock' ) {"True"} else {"False"})]"

                    
                    # padodam: lokālā diska tmp mapi, msi pakotnes izmēru
                    $parametersCheck = @{
                        Session      = $RemoteSession
                        ScriptBlock  = $CheckRemoteSpace
                        ArgumentList = $WorkingDirectory, $InstallFile.Length
                        ErrorAction  = 'Stop'
                    }
                    try {
                    
                        $result = Invoke-Command @parametersCheck
                    }
                    catch {
                        Write-host $_
                        $LogMessage += $_ | Out-String | ForEach-Object { @( "[CheckDiskSize] $_" ) }
                    }

                    Write-Verbose "[Install-CompProgram]:result:[$result]"

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

                        Write-Verbose "[Install-CompProgram]:InstallFile.Name:[$($InstallFile.Name)], WorkingDirectory:[$WorkingDirectory]"
                        # Write-Host "[Install-CompProgram]:[Invoke-VSkInstallBlock] module is available [$(if ( ( Get-Command -Name 'Invoke-VSkInstallBlock' ).name -like 'Invoke-VSkInstallBlock' ) {"True"} else {"False"})]"
                        # padodam: lokālā diska tmp mapi, msi pakotnes datnes nosaukumu

                        # $FileName = "$($InstallFile.Name)"

                        $parametersInstall = @{
                            Session      = $RemoteSession
                            ScriptBlock  = $InstallBlock
                            ArgumentList = $WorkingDirectory, $InstallFile.Name
                            ErrorAction  = 'Stop'
                        }
                        try {

                            $InstallResult = Invoke-Command @parametersInstall

                            $InstallResult |  ForEach-Object { $LogMessage += @( "$_") }
                        }
                        catch {
                            Write-Verbose "[Install-CompProgram] Error: $_"
                            $LogMessage += $_ | Out-String | ForEach-Object { @( "[Install-CompProgram] $_" ) }
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