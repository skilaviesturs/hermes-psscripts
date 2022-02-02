#Requires -Version 5.1
Function Get-CompEvent {
	<#
	.SYNOPSIS
	Skripts pārbauda datora Eventlog uz Windows Update servisa notikumiem
	
	.DESCRIPTION
	Skripts pārbauda datora Eventlog uz Windows Update servisa notikumiem
	
	.PARAMETER Name
	Veic konkrētā datora pārbaudi uz windows jauninājumiem. Rezultātu kopsavilkumu izvada ekrānā. Atbalsta parametru ievadi pipeline.
	
	.PARAMETER InPath
	Veic sarakstā norādīto datoru pārbaudi uz windows jauninājumiem. Rezultātu kopsavilkumu izvada ekrānā.
	
	.PARAMETER Days
	Norādam par kādu periodu pagātnē tiek skatīti notikumi. Pēc noklusējuma - 1 diena.
	
	.PARAMETER Help
	Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.
	
	.EXAMPLE
	Get-WinUpdateStatuss.ps1 -Name EX00001
	Pārbauda datora EX00001 notikumu žurnālu. Rāda tikai kopsavilkumu.
	
	.EXAMPLE
	Get-WinUpdateStatuss.ps1 -InPath EX00001 .\computers.txt -Details
	Sagatavo .\computers.txt norādītajiem datoru notikumu žurnāla ierakstus. Parāda detalizētu atskaiti.
	
	.EXAMPLE
	'EX00001' | Get-WinUpdateStatuss.ps1 -Days 7
	Pārbauda datora EX00001 notikumu žurnālu par notikumiem pēdējās 7 dienās
	
	.NOTES
		Author:	Viesturs Skila
		Version: 1.2.12
	#>
	[CmdletBinding(DefaultParameterSetName = 'Name')]
	param (
		[Parameter(Position = 0, Mandatory,
			ValueFromPipeline,
			ParameterSetName = 'Name',
			HelpMessage = "Name of computer")]
		[ValidateNotNullOrEmpty()]
		[string[]]$Name,
		
		[Parameter(Position = 0, Mandatory = $true,
			ParameterSetName = 'InPath',
			HelpMessage = "Path of txt file with list of computers.")]
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
		[System.IO.FileInfo]$InPath,

		[Parameter(Mandatory = $false, ParameterSetName = 'Name')]
		[Parameter(Mandatory = $false, ParameterSetName = 'InPath')]
		[int]$Days = 30,
		[Parameter(Position = 0, Mandatory = $true,
			ParameterSetName = 'Help'
		)]
		[switch]$Help
	)
	BEGIN {
		<# ---------------------------------------------------------------------------------------------------------
		Skritpa konfigurācijas datnes
		--------------------------------------------------------------------------------------------------------- #>
		$CurVersion = "1.2.12"
		$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()
		$__ScriptName = $MyInvocation.MyCommand
		$__ScriptPath = Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent

		if ($Help) {
			Write-Host "`nVersion:[$CurVersion]`n"
			$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
			$text | ForEach-Object { Write-Host $($_) }
			Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
			Exit
		}
		Function Write-msg { 
			[Alias("wrlog")]
			Param(
				[Parameter(Mandatory = $true)]
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
					Out-File "$LogFile.log" -Append -ErrorAction Stop
				}
				elseif ( $log ) {
					Write-Verbose "[$flag] $text"
					Write-Output "$timeStamp [$ScriptRandomID] [$ScriptUser] [$flag] $text" |
					Out-File "$LogFile.log" -Append -ErrorAction Stop
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
		function Get-PendingUpdates {
			try {
				$UpdateResults = @()
				$updatesession = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session", $Computername))
				$UpdateSearcher = $updatesession.CreateUpdateSearcher()
				$searchresult = $updatesearcher.Search("IsInstalled=0") # 0 = NotInstalled | 1 = Installed
				$Script:ThereIsUpdates = $searchresult.Updates.Count
				if ( $Script:ThereIsUpdates -gt 0 ) {
					#Updates are waiting to be installed
					Write-msg -log -text "$Computername | summary | [$Script:ThereIsUpdates] $(
						if ( $Script:ThereIsUpdates -gt 1 ) {"are"} else {"is"} ) waiting to be installed"   
					foreach ( $entry in $searchresult.Updates ) {
						Write-msg -log -text "$Computername | update | KB$($entry.KBArticleIDs) Title:$($entry.Title)"
						$UpdateResults += @("$Computername | KB$($entry.KBArticleIDs) | Title:$($entry.Title)")
					}
				}
				else {
					Write-msg -log -text "$Computername | update | updates have not been found"
					$UpdateResults = "$Computername | update | updates have not been found"
				}
			}
			catch {
				Write-msg -log -bug -text "[listPendingUpdates] $($_.Exception.Message)"
			}
			return $UpdateResults
		}
		<# ---------------------------------------------------------------------------------------------------------
		Definējam konstantes
		--------------------------------------------------------------------------------------------------------- #>
		# $Out2File = Get-ChildItem -Path $InPathFileName -Attributes Archive
		# $PathToFile = "$($Out2File.DirectoryName)\$($Out2File.BaseName).log"
		Write-Verbose "Got parameters to work with: Name[$Name];`nInPath[$InPath];`nOutPath[$OutPath]=>[$(if($OutPath) {"$PathToFile"})];`nDays[$Days]"
		$MaxThreads = 40
		$Output = @()
		$NoResults = @()
		$NameFromPipe = 0
		$maxAge = (Get-Date).Date.AddDays(-$Days)
		# $StatusCode_ReturnValue = @{
		# 	1    = 'Staged   '
		# 	2    = 'Installed'
		# 	4    = 'KBRestart'
		# 	19   = 'Success  '
		# 	20   = 'Error    '
		# 	21   = 'WURestart'
		# 	27   = 'Paused   '
		# 	43   = 'Started  '
		# 	44   = 'Download '
		# 	1074	= 'Rebooted '
		# 	6005	= 'EventLog '
		# 	6006	= 'EventLog '
		# 	6013	= 'Uptime   '
		# }
		# $statusFriendlyText = @{
		# 	Name       = 'Status'
		# 	Expression = { 
		# 		if ( $null -eq $_.EventID) {
		# 			"N/A"
		# 		}
		# 		else {
		# 			$StatusCode_ReturnValue[([int]$_.EventID)]
		# 		}
		# 	}
		# }
	
		<# Microsoft-Windows-Servicing: 1, 2, 4
		# Microsoft-Windows-WindowsUpdateClient: 19, 20, 21, 27, 43,
		# Microsoft-Windows-User32: 1074
		# Microsoft-Windows-EventLog: 6005, 6006, 6013
		# #>
		$listId = @(1, 2, 4, 19, 20, 21, 43, 1074, 6013)
		$Block = {
			#$Computer	= $args[0]
			#$maxAge		= $args[1]
			#$listID		= $args[2]
			try {
				$resultTotal = @()
				$result = @()
				try {
					$result = Invoke-Command -ComputerName $args[0] -ScriptBlock {
						Param ( $maxAge, $listID )
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
					} -ArgumentList ( $args[1], $args[2] ) -ErrorAction Stop
					if ( $result.count -gt 0 ) { $resultTotal += @($result); $result = @() }
				}
				catch {
	
				}
				try {
					$result = Invoke-Command -ComputerName $args[0] -ScriptBlock {
						Param ( $maxAge, $listID )
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
					} -ArgumentList ( $args[1], $args[2] )
					if ( $result.count -gt 0 ) { $resultTotal += @($result); $result = @() }
				}
				catch {
	
				}
				try {
					$result = Invoke-Command -ComputerName $args[0] -ScriptBlock {
						Param ( $maxAge, $listID )
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
					} -ArgumentList ( $args[1], $args[2] )
					if ( $result.count -gt 0 ) { $resultTotal += @($result); $result = @() }
				}
				catch {
	
				}
				try {
					$result = Invoke-Command -ComputerName $args[0] -ScriptBlock {
						Param ( $maxAge, $listID )
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
					} -ArgumentList ( $args[1], $args[2] )
					if ( $result.count -gt 0 ) { $resultTotal += @($result); $result = @() }
				}
				catch {
	
				}
				$resultTotal
			}
			catch {
				Write-Host "`n$_" -ForegroundColor Yellow
			}
		}
		<# ---------------------------------------------------------------------------------------------------------
		Sākam darbu
		--------------------------------------------------------------------------------------------------------- #>
		Get-Job | Remove-Job
		if ($PSCmdlet.ParameterSetName -eq 'InPath') {
			$Name = Get-Content -Path $InPath | 
			Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  | 
			Sort-Object | Get-Unique
		}#endif
		$JobWatch = [System.Diagnostics.Stopwatch]::startNew()
		Write-Host "[Get-CompEvents]:got:[$($Name.Count)]"
		Write-msg -log -text "[Get-CompEvents]:got:[$($Name.Count)]"
		Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
	}
	
	PROCESS {
		if ($PSCmdlet.ParameterSetName -eq 'InPath') {
			foreach ( $item in $Name ) {
				Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
				While ($(Get-Job -state running).count -ge $MaxThreads) {
					Start-Sleep -Milliseconds 10
				}
				$null = Start-Job -Name "$($item)" -Scriptblock $Block -ArgumentList "$($item)", $maxAge, $listID
			}
		}
		else {
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
			$NameFromPipe++
			While ($(Get-Job -state running).count -ge $MaxThreads) {
				Start-Sleep -Milliseconds 10
			}
			$null = Start-Job -Name "$($Name)" -Scriptblock $Block -ArgumentList "$($Name)", $maxAge, $listID 
		}
	}
	
	END {
		#Write-Host -NoNewLine " done." -ForegroundColor Yellow -BackgroundColor Black
		$OutputTotal = @()
		#Write-Host -NoNewLine "`nRunning      : " -ForegroundColor Yellow -BackgroundColor Black
		While (Get-Job -State "Running") {
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep 5
		}
		#Get information from each job.
		foreach ( $job in Get-Job ) {

			$result = @( Receive-Job -Id ($job.Id) )
			if ( $result -or $result.count -gt 0 ) {
				$result | Add-Member -MemberType NoteProperty -Name 'Computer' -Value "$($job.Name)"
				if ( -not $result.EventID ) { $result.EventID = $result.InstanceId }
				$Output += $result
			}
			else {
				$OutputTotal += New-Object -TypeName psobject -Property @{
					Status   = "Unknown";
					Name     = "$($job.Name)";
					Comments = "No update events found or update not started yet"
				}
				$NoResults += "$($job.Name)"
			}
		}
		#Remove all jobs created.
		Get-Job | Remove-Job
		Stop-Watch -Timer $JobWatch -Name JOBers
		Write-Host "`n================================================================================================="
		Write-Host "Total [$(if( $NameFromPipe -eq 0 ) {"$($name.Count)"}else{"$NameFromPipe"})] computers:"
		Write-host "Got events from [$(($output.Computer | Get-Unique).Count)] computers and no events from [$($NoResults.Count)] computers."

		Stop-Watch -Timer $scriptWatch -Name CompEvents
		
		$Output
	}
}