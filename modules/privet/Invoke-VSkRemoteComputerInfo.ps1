#Requires -Version 5.1
Function Invoke-VSkRemoteComputerInfo {
    param (
        [string]$Computer
    )
    Write-Verbose "[Test_VSkRemote] starting"
    try {

        $ObjectReturn = New-Object -TypeName psobject -Property @{
            Computer    = $Computer
            PingTest    = $Null
            Version     = $Null
            WUModule    = $Null
            Language    = $Null
            Policy      = $Null
            msgCatchErr = [System.Collections.ArrayList]@()
        }

        if ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction Stop ) {
            [string]$ObjectReturn.PingTest = "Success"
            $RemoteSession = New-PSSession -ComputerName $Computer -ErrorAction Stop
            
            [string]$ObjectReturn.Version = Invoke-Command -Session $RemoteSession -ScriptBlock { 
                $PSVersionTable.PSVersion.Major } 
    
            #check module 'PSWindowsUpdate' is installed, if not, copy from script's root directory to remote computer
            [string]$WUModule = ( Invoke-Command -Session $RemoteSession -ScriptBlock { 
                    Get-Module -ListAvailable -Name 'PSWindowsUpdate' } ).Name
                
            if ( $WUModule ) {
                $ObjectReturn.WUModule = $WUModule
            }
            else {
                $parameters = @{
                    Destination = "C:\Program Files\WindowsPowerShell\Modules\"
                    ToSession   = $RemoteSession
                    Recurse     = $True
                    ErrorAction = 'Stop'
                }
                try {
                    Copy-Item "lib\modules\PSWindowsUpdate\" @parameters
                    [string]$ObjectReturn.WUModule = 'PSWindowsUpdate'
                }
                catch {
                    $ObjectReturn.msgCatchErr.Add("[WUModule] Unable copy [PSWindowsUpdate] module to [$Computer]")
                    $ObjectReturn.msgCatchErr.Add("[WUModule] Error: $_")
                }
            }

            #check computer has ExecutionPolicy = RemoteSigned
            [string]$ObjectReturn.Language = [string]( Invoke-Command -Session $RemoteSession -ScriptBlock {
                    $ExecutionContext.SessionState.LanguageMode }
            ).value

            #check computer has LanguageMode = FullLanguage
            [string]$ObjectReturn.Policy = Invoke-Command -Session $RemoteSession -ScriptBlock { Get-ExecutionPolicy }
            
        }
        else {
            $ObjectReturn.msgCatchErr.Add("[PingTest] cannot ping host [$Computer]")
        }
    }
    catch {
        $ObjectReturn.msgCatchErr.Add("[msgCatchErr] $_")
    }
    if ( $RemoteSession.count -gt 0 ) {
        Remove-PSSession -Session $RemoteSession
    }
    $ObjectReturn
}#endblock
