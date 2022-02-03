#Requires -Version 5.1
Function Invoke-VSkEventBlock {
    param (
        $Computer, $maxAge, $listID
    )

    try {

        $RemoteSession = New-PSSession -ComputerName $Computer -ErrorAction Stop

        $resultTotal = [System.Collections.ArrayList]@()

        try {
            $resultServicing = @(
                    Invoke-Command -Session $RemoteSession -ScriptBlock {
                    param ( $maxAge, $listID )

                    Get-WinEvent -FilterHashtable @{
                        Logname      = 'Setup'
                        ProviderName	= 'Microsoft-Windows-Servicing'
                    } -ErrorAction SilentlyContinue | 
                    Where-Object { $_.id -in $listID } | Where-Object { $_.TimeCreated -gt $maxAge } |
                    Select-Object -Property @{Name = 'TimeGenerated'; Expression = { $_.TimeCreated } },
                    MachineName, Source,
                    @{Name = 'EventID'; Expression = { $_.Id } },
                    @{Name = 'KB'; Expression = { if ( $_.Message -match "(KB\d{7})" ) { $matches[1] } } },
                    Message, @{Name = 'EventSource'; Expression = { 'Setup' } }
                
                } -ArgumentList $maxAge, $listID -ErrorAction Stop
            )
        }
        catch {}
        try {
            $resultWindowsUpdateClient = @(
                Invoke-Command -Session $RemoteSession -ScriptBlock {
                    param ( $maxAge, $listID )

                    Get-WinEvent -FilterHashtable @{
                        logname      = 'System'
                        ProviderName	= 'Microsoft-Windows-WindowsUpdateClient'
                    } -ErrorAction SilentlyContinue |
                    Where-Object { $_.id -in $listId } | Where-Object { $_.TimeCreated -gt $maxAge } |
                    Select-Object -Property @{Name = 'TimeGenerated'; Expression = { $_.TimeCreated } },
                    MachineName, Source,
                    @{Name = 'EventID'; Expression = { $_.Id } },
                    @{Name = 'KB'; Expression = { if ( $_.Message -match "(KB\d{7})" ) { $matches[1] } } },
                    Message, @{Name = 'EventSource'; Expression = { 'WUClient' } }
                
                } -ArgumentList $maxAge, $listID -ErrorAction Stop
            )
        }
        catch {}
        try {
            $resultUser32 = @(
                    Invoke-Command -Session $RemoteSession -ScriptBlock {
                    param ( $maxAge, $listID )
                    
                    Get-WinEvent -FilterHashtable @{
                        logname      = 'System'
                        ProviderName	= 'User32'
                    } -ErrorAction SilentlyContinue |
                    Where-Object { $_.id -in $listId } | Where-Object { $_.TimeCreated -gt $maxAge } |
                    Select-Object -Property @{Name = 'TimeGenerated'; Expression = { $_.TimeCreated } },
                    MachineName, Source,
                    @{Name = 'EventID'; Expression = { $_.Id } },
                    @{Name = 'KB'; Expression = { if ( $_.Message -match "(KB\d{7})" ) { $matches[1] } } },
                    Message, @{Name = 'EventSource'; Expression = { 'System' } }
                
                } -ArgumentList $maxAge, $listID -ErrorAction Stop
            )
        }
        catch {}
        try {
            $resultEventLog = @(
                Invoke-Command -Session $RemoteSession -ScriptBlock {
                    param ( $maxAge, $listID )

                    Get-WinEvent -FilterHashtable @{
                        logname      = 'System'
                        ProviderName	= 'EventLog'
                    } -ErrorAction SilentlyContinue |
                    Where-Object { $_.id -in $listId } | Where-Object { $_.TimeCreated -gt $maxAge } |
                    Select-Object -Property @{Name = 'TimeGenerated'; Expression = { $_.TimeCreated } },
                    MachineName, Source,
                    @{Name = 'EventID'; Expression = { $_.Id } },
                    @{Name = 'KB'; Expression = { if ( $_.Message -match "(KB\d{7})" ) { $matches[1] } } },
                    Message, @{Name = 'EventSource'; Expression = { 'System' } }
                
                } -ArgumentList $maxAge, $listID -ErrorAction Stop
            )
        }
        catch {}
    }
    catch{
        Write-Host "[Get-CompEvent] cannot connect to [$Computer]"
    }
    finally {
        if ( $RemoteSession.Count -gt 0 ) {
            Remove-PSSession -Session $RemoteSession
        }
    }

    if ($resultServicing.TimeGenerated.Count -gt 0 ) {
        $resultTotal += $resultServicing
    }
    if ($resultWindowsUpdateClient.TimeGenerated.Count -gt 0 ) {
        $resultTotal += $resultWindowsUpdateClient
    }
    if ($resultUser32.TimeGenerated.Count -gt 0 ) {
        $resultTotal += $resultUser32; 
    }
    if ($resultEventLog.TimeGenerated.Count -gt 0 ) {
        $resultTotal += $resultEventLog
    }

    $resultTotal
}