#Requires -Version 5.1
Function Invoke-CompWakeOnLan {
    <#
    .SYNOPSIS
    Skripts datu bāzē atrod nepieciešamos parametrus datora attālinātai sāknēšanai
    
    .PARAMETER ComputerName
    Datora vārds
    
    .PARAMETER DataArchiveFile
    Datu bāzes faila atrašanās vieta
    
    .NOTES
        Author:	Viesturs Skila
        Version: 1.1.1
    #>
    [CmdletBinding(DefaultParameterSetName = 'Name')]
    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline, ParameterSetName = 'Name',
            HelpMessage = "Name of computer"
        )]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName,
    
        [Parameter(Position = 1, Mandatory = $true, ParameterSetName = 'Name',
            HelpMessage = "Path to archive file."
        )]
        [System.IO.FileInfo]$DataArchiveFile,
    
        # [Parameter(Position = 2, Mandatory = $true, ParameterSetName = 'Name',
        #     HelpMessage = "Path to archive file."
        # )]
        # [System.IO.FileInfo]$CompTestOnlineFile,
    
        [Parameter(Position = 0, Mandatory = $true,
            ParameterSetName = 'Help'
        )]
        [switch]$Help
    )
    BEGIN {
        <# ---------------------------------------------------------------------------------------------------------
        Skripta konfigurācijas datnes
        --------------------------------------------------------------------------------------------------------- #>
        $CurVersion = "1.1.1"
        $scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
    
        $LogObject = @()
    
        if ( Test-Path $DataArchiveFile -PathType Leaf ) {
            try {
                $DataArchive = @(Import-Clixml -Path $DataArchiveFile -ErrorAction Stop)
            }
            catch {
                $DataArchive = @()
            }
        }
        else {
            $DataArchive = @()
        }
        #$LogObject += @("[Waker] [INFO] got:[$($ComputerName.Count)]")
        #$LogObject += @("[Waker] [INFO] got DataArchive:[$($DataArchive.Count)]")
    }
    
    PROCESS {
        [string]$HostDNSName = $null
        [string]$TargetMacAddress = $null
        try {
            [string]$IPAddress = [System.Net.Dns]::GetHostAddresses($ComputerName)
            [string]$Mask = $IPAddress -match '^([0-9]{1,3}[.]){2}([0-9]{1,3})'
        }
        catch {
            $LogObject += @("[Waker] [ERROR] Computer [$ComputerName] is not registered in DNS. Exit.")
        }
        #Atrodam IP adreses segmentu xxx.xxx.xxx
        if ( $Mask ) {
            [string]$Pattern = $matches[0]
            #$LogObject += @("[Waker] [INFO] [$ComputerName] belongs to net segment [$Pattern]")
            if ( $DataArchive.Count -gt 0 ) {
                #Atrodam arhīvā mērķa datora ierakstu, lai noteiktu mac adresi
                foreach ( $rec in $DataArchive ) {
                    if ( $rec.DNSName -like $ComputerName -or $rec.PipedName -like $ComputerName ) {
                        $TargetMacAddress = $rec.MacAddress
                        $LogObject += @("[Waker] [INFO] [$ComputerName] has mac address [$TargetMacAddress]")
                        break
                    }
                }
                #Atrodam arhīvā datoru, kas atrodas tajā pašā segmentā, lai no tā varētu pamodināt guļošo
                foreach ( $rec in $DataArchive ) {
                    if ( $rec.IPAddress -match "$Pattern`*" -and (
                         $rec.DNSName -notlike $ComputerName -or
                         $rec.PipedName -notlike $ComputerName ) ) {

                        # $OnlineRemoteComps = Invoke-Expression "& `"$CompTestOnlineFile`" `-Name $($rec.DNSName) "
                        $OnlineRemoteComps = Invoke-Command -ScriptBlock ${Function:Get-CompTestOnline} -ArgumentList $rec.DNSName
                        #Pārbaudam vai atrastā remote datora WinRM serviss darbojas
                        if ( $OnlineRemoteComps.WinRMservice ) {
                            $HostDNSName = $rec.DNSName
                            $LogObject += @("[Waker] [INFO] found online neighbor [$HostDNSName]:[$($rec.IPAddress)] on the same net.")
                            break
                        }
                    }
                }
                if ( [string]::IsNullOrWhitespace($HostDNSName) ) {
                    $LogObject += @("[Waker] [ERROR] there is no entry for the computer on the same net [$Pattern] in the database. Exit.")
                }
                elseif ( [string]::IsNullOrWhitespace($TargetMacAddress) ) {
                    $LogObject += @("[Waker] [ERROR] there is no entry for the computer [$ComputerName] mac address in the database. Exit.")
                }
                else {
                    $LogObject += @("[Waker] [INFO] going to WakeOnLan [$TargetMacAddress] from remote host [$HostDNSName].")
                    $result = Invoke-Command -Computername $HostDNSName -ScriptBlock ${Function:Invoke-WakeOnLan} -ArgumentList $TargetMacAddress
                    $result | ForEach-Object { $LogObject += @($_) }
                }
            }
        }
    
    }
    
    END {
        $result = Stop-Watch -Timer $scriptWatch -Name Script
        $result | ForEach-Object { $LogObject += @($_) }
        $Output = @()
        $i = 0
        $LogObject | ForEach-Object {
            $Output += @( New-Object -TypeName psobject -Property @{
                    id       = $i;
                    Computer = [string]$ComputerName;
                    Message  = [string]$_;
                }
            )
            $i++
        }
    
        $Output
    }
}