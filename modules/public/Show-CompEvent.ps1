#Requires -Version 5.1
Function Show-CompEvent {
	<#
	.SYNOPSIS
	Skripts transformē un attēlo uz ekrāna no Get-CompEvent saņemto objectu
	
	.DESCRIPTION
	Skripts transformē un attēlo uz ekrāna no Get-CompEvent saņemto objectu

    .PARAMETER InputObject
	Norāda Get-CompEvent izveidoto objektu.
	
    .PARAMETER OutPath
	Norāda datnes vārdu, kurā tiks ierakstīts skripta rezultāta kopsavilkums.
    Ja parametrs nav norādīts, rezultāts tiek izvadīts uz ekrāna.

    .PARAMETER InPathFileName
	Rezultātu kopsavilkumu ieakstīšana norādītājā datnē.
	
	.NOTES
		Author:	Viesturs Skila
		Version: 1.0.1
	#>
	[CmdletBinding()]
	param (
        [Parameter(Position = 0, Mandatory)]
        [object]$InputObject,

        [switch]$Named,
        
		[switch]$OutPath,
		[ValidateScript( {
				if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
					Write-Host "File does not exist"
					throw
				}#endif
				if ( $_ -notmatch ".txt|.tmp") {
					Write-Host "The file specified in the path argument must be text file"
					throw
				}#endif
				return $true
			} ) ]
		[System.IO.FileInfo]$InPathFileName
    )
        
    # $LogFileDir = "log"
    # $LogFile = "$LogFileDir\RemoteJob_$(Get-Date -Format "yyyyMMdd")"
    $Out2File = Get-ChildItem -Path $InPathFileName -Attributes Archive
	$PathToFile = "$($Out2File.DirectoryName)\$($Out2File.BaseName).log"

    $StatusCode_ReturnValue = @{
        1    = 'Staged   '
        2    = 'Installed'
        4    = 'KBRestart'
        19   = 'Success  '
        20   = 'Error    '
        21   = 'WURestart'
        27   = 'Paused   '
        43   = 'Started  '
        44   = 'Download '
        1074	= 'Rebooted '
        6005	= 'EventLog '
        6006	= 'EventLog '
        6013	= 'Uptime   '
    }

    $statusFriendlyText = @{
        Name       = 'Status'
        Expression = { 
            if ( $null -eq $_.EventID) {
                "N/A"
            }
            else {
                $StatusCode_ReturnValue[([int]$_.EventID)]
            }
        }
    }
    <# ---------------------------------------------------------------------------------------------------------
        Sagatavojam kopsavilkumu
    --------------------------------------------------------------------------------------------------------- #>
    $WindowUpdateList = @()
    $ComputersInvolved = @()
    $OutputTotal = @()

    $InputObject | ForEach-Object {
        if ( $ComputersInvolved -notcontains "$($_.Computer)" ) {
            # Write-Host "$($_.Computer)"
            $ComputersInvolved += @($_.Computer)
        }
    }

    # $ComputersInvolved.GetType().BaseType.name
    # $InputObject.GetType().BaseType.name

    $ComputersInvolved | ForEach-Object {
        [int]$CompError = 0
        [int]$Success = 0
        [int]$Started = 0
        [int]$RestartRequired	= 0
        [string]$ErrorMsg = ''
        foreach ( $row in $InputObject ) {
            if ( $_ -like $row.Computer ) {
                switch ($row.EventID) {
                    19	{ $Success++ }
                    20	{ $CompError++; $ErrorMsg = (($row.Message.Split(':'))[2]).trim(); }
                    21	{ $RestartRequired++ }
                    #27	{ $AutoUpdatePaused++ }
                    43	{ $Started++ }
                }
            }
        }
        if ( $CompError -gt 0 ) {
            $OutputTotal += New-Object -TypeName psobject -Property @{
                Status   = "Error" ;
                Name     = $_ ;
                Comments = $ErrorMsg ;
            }
        }
        if ( $RestartRequired -gt 0 ) {
            $OutputTotal += New-Object -TypeName psobject -Property @{
                Status   = "RestartRequired" ;
                Name     = $_ ;
                Comments = $ErrorMsg ;
            }
        }
        else {
            if ( $Success -eq $Started -and $Started -ne 0 -and $Success -ne 0 ) {
                $OutputTotal += New-Object -TypeName psobject -Property @{
                    Status   = "Successfull" ;
                    Name     = $_ ;
                    Comments = "Started[$Started]=>Success[$Success]: done." ;
                }
            }
            if ( $Success -gt $Started ) {
                $OutputTotal += New-Object -TypeName psobject -Property @{
                    Status   = "Success" ;
                    Name     = $_ ;
                    Comments = "Started[$Started]=>Success[$Success]: done, but check logs for sure." ;
                }
            }
            if ( $Success -lt $Started ) {
                $OutputTotal += New-Object -TypeName psobject -Property @{
                    Status   = "Updating" ;
                    Name     = $_ ;
                    Comments = "Started[$Started]=>Success[$Success]: check logs." ;
                }
            }
        }
    }
    #Sagatavojam uzstādīto jauninājumu sarakstu
    ForEach ( $row in $InputObject ) {
        if ( $row.EventID -eq 19 ) {
            $msg = (($row.Message.Split(':'))[2]).trim()
            if ( $WindowUpdateList.contains($msg) -eq $False ) {
                $WindowUpdateList += $msg
            }
        }
    }

    <# ---------------------------------------------------------------------------------------------------------
        Attēlojam kopsavilkumu atbilstoši [-Detailed] un [-OutPath] statusiem
    --------------------------------------------------------------------------------------------------------- #>
    #ja OutPath nav iestatīts
    if ( $OutPath ) {
        Write-Output "`n[$(Get-Date -Format "yyyy.MM.dd HH:mm:ss")]------------------------------------------------------------------------------------------" | 
        Out-File -FilePath $PathToFile -Encoding ASCII -Force
        #Windows jauninājumu atskaite failā
        Write-Output "`nSuccessfully installed Windows updates:`n=======================================" | 
        Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        $WindowUpdateList | Out-String -Stream | Where-Object { $_ -ne "" } | Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        
        #Kopsavilkuma atskaite failā
        Write-Output "`nStatuss of updates:`n===================" | Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        $OutputTotal | Sort-Object -Property Status | Format-Table Status, Name, Comments -AutoSize | Out-String -Stream | 
        Where-Object { $_ -ne "" } | Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        
        #Avota informācija failā
        Write-Output "`nFrom computer's event log:`n==========================" | Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        foreach ( $computer in $ComputersInvolved ) {
            Write-Output "`nComputer:[$Computer]====================================================" | 
            Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
            $InputObject | Sort-Object -Property TimeGenerated -Descending | Where-Object -Property Computer -like $computer `
            | Select-Object TimeGenerated, Computer, EventSource, $statusFriendlyText, KB, Message `
            | Format-Table TimeGenerated, EventSource, Status, KB, Message -AutoSize  `
            | Out-String -Width 1024 -Stream | Where-Object { $_ -ne "" } `
            | Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force
        }
        Write-Output "---------------------------------------------------------------------------------------------------------------" | 
        Out-File -FilePath $PathToFile -Encoding ASCII -Append -Force	
    }
    #Avota informācija ekrānā
    if ($Named) {
        Write-Host "`nFrom computer's event log:"
        Write-Host "=========================="
        $InputObject | Sort-Object -Property TimeGenerated -Descending | Select-Object TimeGenerated, EventSource, $statusFriendlyText, KB, Message `
        | Format-Table * -AutoSize  `
        | Out-String -Stream | Where-Object { $_ -ne "" } `
        | ForEach-Object { `
                if ($_.Contains('Error')) { Write-Host "$_" -ForegroundColor Red } 
            elseif ( $_.Contains('WURestart') -or $_.Contains('KBRestart') ) { Write-Host "$_" -ForegroundColor Yellow } 
            else { Write-Host "$_" } 
        }
    }
    #Kopsavilkuma atskaite ekrānā
    Write-Host "`nSuccessfully installed Windows updates:"
    Write-Host "======================================="
    $WindowUpdateList | Format-Table * -AutoSize | Out-String -Stream | Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" } 
    Write-Host "`nStatuss of updates:"
    Write-Host "==================="
    $OutputTotal | Sort-Object -Property Status | Format-Table Status, Name, Comments -AutoSize | Out-String -Stream | Where-Object { $_ -ne "" } `
    | ForEach-Object { `
            if ($_.Contains('Error')) { Write-Host "$_" -ForegroundColor Red } 
        elseif ( $_.Contains('WURestart') -or $_.Contains('KBRestart') ) { Write-Host "$_" -ForegroundColor Yellow } 
        else { Write-Host "$_" } 
    }
    if ( $OutPath ) { Write-Host "`nThe report file is [$PathToFile]." -ForegroundColor Yellow }

    # Stop-Watch -Timer $scriptWatch -Name CompEvents
}