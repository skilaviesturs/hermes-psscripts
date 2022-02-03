#Requires -Version 5.1
Function Invoke-VSkInstallBlock {
    param( $tempPath, $FileName )
    try {
        $LogMessage = [System.Collections.ArrayList]@()
        # $LogMessage = @()
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

            $SourceFile = Get-ChildItem -Path "$tempPath\$FileName"
            Write-Host "[InstallBlock] [$tempPath]\[$FileName]" 
            Write-Host "[InstallBlock] SourceFile:[$($SourceFile.FullName)]" 

            if ( $SourceFile.Extension -eq '.msi' ) {

                $DataStamp = get-date -Format yyyyMMddTHHmmss
                $logFile = '{0}-{1}.log' -f $SourceFile.fullname, $DataStamp
                $MSIArguments = @(
                    "/i"
                    ('"{0}"' -f $SourceFile.fullname)
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
            elseif ( $SourceFile.Extension -eq '.exe'  ) {

                $logFile = "$(Split-Path -Path "$($SourceFile.FullName)" -Parent)\$($SourceFile.BaseName).log"
                # $logFile = '{0}-{1}.log' -f $SourceFile.FullName, $DataStamp
                $LogMessage += @("[Installer] [INFO] [$FileName] log file [$logFile]")

                if ( $SourceFile.BaseName -like "AcroRdrDC*"  ) {
                    $Arguments = "`/c $($SourceFile.FullName) `/sAll /rs /msi EULA_ACCEPT=YES /L*V $logFile"
                }
                else {
                    $Arguments = "`/c $($SourceFile.FullName) `/S /L*V $logFile"
                }
                
                $object = New-object System.Diagnostics.ProcessStartInfo -Property @{
                    CreateNoWindow         = $true
                    UseShellExecute        = $false
                    RedirectStandardOutput = $true
                    RedirectStandardError  = $true
                    FileName               = 'cmd.exe'
                    Arguments              = $Arguments
                    # WorkingDirectory       = $tempPath
                    WorkingDirectory       = "$(Split-Path -Path "$($SourceFile.FullName)" -Parent)"
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
        $LogMessage += $_ | Out-String | ForEach-Object { @( "[Installer] Error: $_" ) }
    }
    finally {
        if ( Test-Path -Path $SourceFile -PathType Leaf ) {
            $SourceFile | Remove-Item -Force
        }
    }
    
    $LogMessage
}