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
    
    .PARAMETER DisplayName
    Norādam uzstādāmās programmatūras DisplayName.
    Programmatūras identifikatoru varam iegūt ar skriptu Get-CompSoftware.ps1
    
    .PARAMETER Help
    Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.
    
    .EXAMPLE
    Install-CompProgram.ps1 -ComputerName EX00001 -InstallPath 'D:\install\7-zip\7z1900-x64.msi'
    Uzstādam uz datora EX00001 programmatūras instalācijas pakotni 7z1900-x64.msi
    
    .NOTES
        Author:	Viesturs Skila
        Version: 1.2.3
    #>
    [CmdletBinding(DefaultParameterSetName = 'Install')]
    Param(
        [Parameter(Position = 0, Mandatory = $True,
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
    
        [Parameter(Position = 1, Mandatory = $True,
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
        [switch]$Help = $False
    )
    BEGIN {
        <# ---------------------------------------------------------------------------------------------------------
        Skritpa konfigurācijas datnes
        --------------------------------------------------------------------------------------------------------- #>
        $CurVersion = "1.2.3"
        #$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
        $__ScriptName = $MyInvocation.MyCommand
        $__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
        #$LogFileDir = "log"
        $LogMessage = @()

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

        # instalējam programmatūru uz mērķa datoru
        $Install = {
            param(
                [Parameter(Position = 0)]
                [string]$tempPath,
                [Parameter(Position = 1)]
                [string]$FileName
            )
            try {
                $LogMessage = @()
                $LogMessage += @("[Installer] [INFO] got $tempPath\$FileName")
            
                #region sesijas lietotājam iestatam kontroli pār pagaidu direktoriju
                #$tempPath = "C:\temp"

                $packagePath = Get-ChildItem -Path $tempPath -Recurse
                $Acl = Get-Acl -Path $tempPath
                $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$(whoami)", "FullControl", "Allow")
                $Acl.SetAccessRule($AccessRule)
                $packagePath | ForEach-Object { Set-Acl -Path $_.FullName -AclObject $Acl }

                #endregion

                if ( Test-Path -Path "$tempPath\$FileName" -PathType Leaf -ErrorAction Stop ) {

                    $file = Get-ChildItem -Path "$tempPath\$FileName"

                    if ( $file.Extension -eq '.msi' ) {

                        $DataStamp = get-date -Format yyyyMMddTHHmmss
                        $logFile = '{0}-{1}.log' -f $file.fullname, $DataStamp
                        $MSIArguments = @(
                            "/i"
                            ('"{0}"' -f $file.fullname)
                            "/qn"
                            "/norestart"
                            "AGREETOLICENSE=yes"
                            "/L*v"
                            $logFile
                        )
                        $object = Start-Process "msiexec.exe" -ArgumentList $MSIArguments -Wait -NoNewWindow -PassThru
                        $LogMessage += @("[Installer] [INFO] [$FileName] log file [$logFile]")

                        #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                        $patterns = @(
                            '-- Installation completed successfully.'
                            'Reconfiguration success or error status: 0.'
                            'Installation success or error status: 0.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[Installer] [SUCCESS] $output")
                            }
                        }
                        $patterns = @(
                            'Windows Installer requires a system restart.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[Installer] [WARN] $output")
                            }
                        }
                    }
                    elseif ( $file.Extension -eq '.exe'  ) {

                        $logFile = "$(Split-Path -Path "$($file.fullname)" -Parent)\$($file.BaseName).log"
                        $LogMessage += @("[Installer] [INFO] [$FileName] log file [$logFile]")

                        if ( $file.BaseName -like "AcroRdrDC*"  ) {
                            $Arguments = "`/c $($file.FullName) `/sAll /rs /msi EULA_ACCEPT=YES /L*V $logFile"
                        }
                        else {
                            $Arguments = "`/c $($file.FullName) `/S /L*V $logFile"
                        }
                        
                        $object = New-object System.Diagnostics.ProcessStartInfo -Property @{
                            CreateNoWindow         = $true
                            UseShellExecute        = $false
                            RedirectStandardOutput = $true
                            RedirectStandardError  = $true
                            FileName               = 'cmd.exe'
                            Arguments              = $Arguments
                            WorkingDirectory       = "$(Split-Path -Path "$($file.fullname)" -Parent)"
                        }
                        $process = New-Object System.Diagnostics.Process 
                        $process.StartInfo = $object 
                        $null = $process.Start()
                        $output = $process.StandardOutput.ReadToEnd()
                        $outputErr = $process.StandardError.ReadToEnd()
                        $process.WaitForExit() 
                        if ( $output ) { $output | Out-File $logFile -Append }
                        if ( $outputErr ) { $outputErr | Out-File $logFile -Append }

                        if ($process.ExitCode -eq 0) { 
                            $LogMessage += @("[Installer] [SUCCESS] process successfull")
                        }
                        else { 
                            $LogMessage += @("[Installer] [ERROR] process failed with error code [$($process.ExitCode)]")
                        }

                        #pārbaudam logfailu uz veiksmīgiem paziņojumiem
                        $patterns = @(
                            '-- Installation completed successfully.'
                            'Reconfiguration success or error status: 0.'
                            'Installation success or error status: 0.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[Installer] [SUCCESS] $output")
                            }
                        }
                        $patterns = @(
                            'Windows Installer requires a system restart.'
                        )
                        $patterns | ForEach-Object {
                            if ( Get-Content $logFile | Select-String -Pattern "$_" ) {
                                $output = Select-String -Path $logFile -Pattern "$_" -CaseSensitive
                                $LogMessage += @("[Installer] [WARN] $output")
                            }
                        }

                    }
                    else {
                        $LogMessage += @("[Installer] [ERROR] supports only msi or exe format")
                    }
                }
                else {
                    $LogMessage += @("[Installer] [ERROR] not found [$tempPath\$FileName]")
                }
            }
            catch {
                $LogMessage += $_ | Out-String | ForEach-Object { @( "$_" ) }
            }
            finally {
                if ( Test-Path -Path $file -PathType Leaf ) {
                    $file | Remove-Item -Force
                }
            }
            
            $LogMessage
        }

    }

    PROCESS {

        try {
            $RemoteSession = New-PSSession -ComputerName $ComputerName -ErrorAction Stop
        
            if ( $RemoteSession.Count -gt 0 ) {

                if ( $PSCmdlet.ParameterSetName -eq "Install" ) {
                    try {
                        $msiFile = Get-ChildItem -Path $InstallPath -ErrorAction Stop
    
                        # padodam: lokālā diska tmp mapi, msi pakotnes izmēru
                        $parameters = @{
                            Session      = $RemoteSession
                            ScriptBlock  = $CheckSpace
                            ArgumentList = "C:\temp", $msiFile.Length
                            ErrorAction  = 'Stop'
                        }

                        $result = Invoke-Command @parameters

                        if ( $result -eq 'Ok' ) {
                            #region Kopējam msi pakotni uz remote datora mapi
                            $parameters = @{
                                Path        = $msiFile.FullName
                                Destination = "C:\temp"
                                ToSession   = $RemoteSession
                                Force       = $true
                                ErrorAction = 'Stop'
                            }
                            Copy-Item @parameters
                            #endregion
    
                            # padodam: lokālā diska tmp mapi, msi pakotnes datnes nosaukumu
                            $parameters = @{
                                Session      = $RemoteSession
                                ScriptBlock  = $Install
                                ArgumentList = "C:\temp", $msiFile.Name
                                ErrorAction  = 'Stop'
                            }
                            $InstallResult = Invoke-Command @parameters 
                            $InstallResult |  ForEach-Object { $LogMessage += @($_) }
                        }
                        else {
                            $LogMessage += @($result)
                        }
                    }
                    catch {
                        $LogMessage += $_ | Out-String | ForEach-Object { @( "$_" ) }
                    }
                }
            }
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

        $Output = @()
        $i = 0
        $LogMessage | ForEach-Object {
            $Output += @( New-Object -TypeName psobject -Property @{
                    id       = $i
                    Computer	= [string]$ComputerName;
                    Message  = [string]$_;
                }#endobject
            )
            $i++
        }

        $Output
    }
}