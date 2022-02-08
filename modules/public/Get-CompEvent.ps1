#Requires -Version 5.1
Function Get-CompEvent {
	<#
	.SYNOPSIS
	Skripts pārbauda datora Eventlog uz Windows Update servisa notikumiem
	
	.DESCRIPTION
	Skripts pārbauda datora Eventlog uz Windows Update servisa notikumiem
	
	.PARAMETER ComputerName
	Veic konkrētā datora pārbaudi uz windows jauninājumiem. Rezultātu kopsavilkumu izvada ekrānā. Atbalsta parametru ievadi pipeline.
	
	.PARAMETER InPath
	Veic sarakstā norādīto datoru pārbaudi uz windows jauninājumiem. Rezultātu kopsavilkumu izvada ekrānā.
	
	.PARAMETER Days
	Norādam par kādu periodu pagātnē tiek skatīti notikumi. Pēc noklusējuma - 1 diena.
	
	.PARAMETER Help
	Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.
	
	.EXAMPLE
	Get-CompEvent.ps1 -ComputerName EX00001
	Pārbauda datora EX00001 notikumu žurnālu. Rāda tikai kopsavilkumu.
	
	.EXAMPLE
	Get-CompEvent.ps1 -InPath EX00001 .\computers.txt -Details
	Sagatavo .\computers.txt norādītajiem datoru notikumu žurnāla ierakstus. Parāda detalizētu atskaiti.
	
	.EXAMPLE
	'EX00001' | Get-CompEvent.ps1 -Days 7
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
		[string[]]$ComputerName,
		
		[Parameter(Position = 0, Mandatory,
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

		[int]$Days = 30,

		[Parameter(Position = 0, Mandatory,
			ParameterSetName = 'Help'
		)]
		[switch]$Help
	)
	BEGIN {
		if (-not $PSBoundParameters.ContainsKey('Verbose')) {
			$VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
		}
		<# ---------------------------------------------------------------------------------------------------------
		Skritpa konfigurācijas datnes
		--------------------------------------------------------------------------------------------------------- #>
		$scriptWatch = [System.Diagnostics.Stopwatch]::startNew()

		<# ---------------------------------------------------------------------------------------------------------
		Definējam konstantes
		--------------------------------------------------------------------------------------------------------- #>
		# $Out2File = Get-ChildItem -Path $InPathFileName -Attributes Archive
		# $PathToFile = "$($Out2File.DirectoryName)\$($Out2File.BaseName).log"
		$MaxThreads = 25
		$maxAge = (Get-Date).Date.AddDays(-$Days)
		Write-Verbose "[Get-CompEvent] ------------------------------------------------------------"
		Write-Verbose "[Get-CompEvent]:Name[$ComputerName]"
		Write-Verbose "[Get-CompEvent]:InPath[$InPath]"
		Write-Verbose "[Get-CompEvent]:Days[$Days], maxAge[$maxAge], MaxThreads[$MaxThreads]"
		$Output = [System.Collections.ArrayList]@()
		$NoResults = @()
		$NameFromPipe = 0
	
		<# Microsoft-Windows-Servicing: 1, 2, 4
		# Microsoft-Windows-WindowsUpdateClient: 19, 20, 21, 27, 43,
		# Microsoft-Windows-User32: 1074
		# Microsoft-Windows-EventLog: 6005, 6006, 6013
		# #>
		$listId = @(1, 2, 4, 19, 20, 21, 43, 1074, 6013)

		<# ---------------------------------------------------------------------------------------------------------
		Sākam darbu
		--------------------------------------------------------------------------------------------------------- #>
		Get-Job | Remove-Job
		if ($PSCmdlet.ParameterSetName -eq 'InPath') {
			$Computers = Get-Content -Path $InPath | 
			Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  | 
			Sort-Object | Get-Unique
		}
		$JobWatch = [System.Diagnostics.Stopwatch]::startNew()
		Write-Verbose "[Get-CompEvent]:got by InPath:[$($Computers.Count)] computers"
		# Write-msg -log -text "[Get-CompEvent]:got:[$($Computers.Count)]"
		Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
	}
	
	PROCESS {
		if ($PSCmdlet.ParameterSetName -eq 'InPath') {
			foreach ( $Computer in $Computers ) {
				Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
				While ($(Get-Job -state running).count -ge $MaxThreads) {
					Start-Sleep -Milliseconds 10
				}
				$null = Start-Job -Name "$Computer" -Scriptblock ${Function:Invoke-VSkEventBlock} -ArgumentList $Computer, $maxAge, $listID
			}
		}
		else {
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
			$NameFromPipe++
			While ($(Get-Job -state running).count -ge $MaxThreads) {
				Start-Sleep -Milliseconds 10
			}
			$null = Start-Job -Name "$ComputerName" -Scriptblock ${Function:Invoke-VSkEventBlock} -ArgumentList $ComputerName, $maxAge, $listID
		}
	}
	
	END {

		$OutputTotal = @()
		While (Get-Job -State "Running") {
			Write-Host -NoNewLine "." -ForegroundColor Yellow -BackgroundColor Black
			Start-Sleep 5
		}
		#Retrieve information from each job.
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

		Get-Job | Remove-Job
		Stop-Watch -Timer $JobWatch -Name JOBers
		# Write-Host "`n================================================================================================="
		Write-Verbose "[Get-CompEvent] overall precessed [$(if( $NameFromPipe -eq 0 ) {"$($Computers.Count)"}else{"$NameFromPipe"})] computers:"
		Write-Verbose "[Get-CompEvent] Got events from [$(($output.Computer | Get-Unique).Count)] computers; no events from [$($NoResults.Count)] computers."

		Stop-Watch -Timer $scriptWatch -Name 'Get-CompEvent'
		
		$Output
	}
}