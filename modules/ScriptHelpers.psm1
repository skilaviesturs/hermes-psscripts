

#region Get-ScriptHelp
Function Get-ScriptHelp {
    <#
    .SYNOPSIS
    Parāda lietotājam skripta on-line help

    .PARAMETER Version
    Norādam skripta versiju

    .PARAMETER ScriptPath
    Norādam skripta atrašanās ceļu

    #>
    param (
        [Parameter(Mandatory)]
        [string] $Version,
        [Parameter(Mandatory)]
        [System.IO.FileInfo] $ScriptPath 
    )

    Write-Host "`nVersion:[$Version]`n"
    $text = Get-Command -Name "$ScriptPath" -Syntax
    $text | ForEach-Object { Write-Host $($_) }
    Write-Host "For more info write `'Get-Help $ScriptPath -Examples`'"
}
#endregion

#region Stop-Watch 
<#
Skripta izpildes taimera output formatēšana
#>
Function Stop-Watch {
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]$Timer,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    $Timer.Stop()
    if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 ) { $bMin = "0$($Timer.Elapsed.Minutes)" } else { $bMin = "$($Timer.Elapsed.Minutes)" }
    if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) { $bSec = "0$($Timer.Elapsed.Seconds)" } else { $bSec = "$($Timer.Elapsed.Seconds)" }
    if ($Name -notlike 'JOBers') {
        Write-msg -log -text "[$Name] finished in $(
            if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
            elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
            else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
        )"
        Write-Host "[$Name] finished in $(
        if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
        elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
        else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
        )"
    }
    else {
        Write-Host "`rJobs done in $(
            if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
            elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
            else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
            )" -ForegroundColor Yellow -BackgroundColor Black
    }
}#endOffunction

#endregion 

#region Write-msg
Function Write-msg { 
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$text,
        [switch]$log,
        [switch]$bug
    )

    try {
        #write-debug "[wrlog] Log path: $log"
        if ( $bug ) { $flag = 'ERROR' } else { $flag = 'INFO' }
        $timeStamp = Get-Date -Format "yyyy.MM.dd HH:mm:ss"
        if ( $log -and $bug ) {
            Write-Warning "[$flag] $text"	
            Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |
                Out-File "$Script:LogFile.log" -Append -ErrorAction Stop
        }
        elseif ( $log ) {
            Write-Verbose "[$flag] $text"
            Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |
                Out-File "$Script:LogFile.log" -Append -ErrorAction Stop
        }
        else {
            Write-Verbose "$flag [$ScriptRandomID] $text"
        }
    }
    catch {
        Write-Warning "[Write-msg] $($_.Exception.Message)"
        return
    }
}
#endregion

#region Write-ErrorMsg
<#
Formatējam kļūdas paziņojumus logfailā
#>
Function Write-ErrorMsg {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]$InputObject,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )
    Write-msg -log -bug -text "[$name] Error: $($InputObject.Exception.Message)"
    $string_err = $InputObject | Out-String
    Write-msg -log -bug -text "$string_err"
}
#endregion

#region Get-ScriptFileUpdate
<#
    Mekējam vai skriptiem nav jaunākas versijas UpdatePath vietnē
#>
Function Get-ScriptFileUpdate {
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$FileName,
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$ScriptPath,
        [Parameter(Position = 2, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$UpdatePath
    )
    $ScriptFile = Get-ChildItem "$ScriptPath\$FileName"  -ErrorAction SilentlyContinue
    $NewFile = Get-ChildItem "$UpdatePath\$FileName" -ErrorAction SilentlyContinue
    if ( $NewFile.count -gt 0 ) {
        if ( $NewFile.LastWriteTime -gt $ScriptFile.LastWriteTime ) {
            Write-msg -log -text "[ScriptUpdate] Found update for script [$FileName]"
            Write-msg -log -text "[ScriptUpdate] Old version [$($ScriptFile.LastWriteTime)], [$($ScriptFile.FullName)]"
            Write-msg -log -text "[ScriptUpdate] New version [$($NewFile.LastWriteTime)], [$($NewFile.FullName)]"
            try {
                Copy-Item -Path $NewFile.FullName -Destination "$(Split-Path -Path "$ScriptPath\$FileName")" -Force -ErrorAction Stop
                Write-msg -log -text "[ScriptUpdate] New version deployed."
            }
            catch {
                #Write-msg -log -bug -text "[ScriptUpdate] [$FileName] $($_.Exception.Message)"
                Write-ErrorMsg -Name 'ScriptUpdate' -InputObject $_
            }
        }
    }

}
#endregion