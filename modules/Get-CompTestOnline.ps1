#Requires -Version 5.1
Function Get-CompTestOnline {
	<#
	.SYNOPSIS
	Skenējam datorvārdus, ip adreses ar ICMP un atgriežam sarakstu ar datoru vārdiem un IP adresēm, kas atbildēja.
	
	.DESCRIPTION
	Skripts skenē norādītos DNS vai IP adreses. Skripts atbalsta parametru padošanu pipeline
	
	.PARAMETER ComputerNames
	Norādam tīkla segmentu bez pēdējā punkta, piemēram, "192.168.0"
	
	.EXAMPLE
	Get-CompTestOnline.ps1 "192.168.0.2"
	Norādam bez parametra 
	
	.EXAMPLE
	Get-CompTestOnline.ps1 -Network "192.168.0.4", "computer.ltb.lan"
	Norādam ar parametru
	
	.EXAMPLE
	Get-Content .\expo-segments.txt | Get-CompTestOnline.ps1
	Padodam parametru no pipeline
	
	.NOTES
	Author:	Viesturs Skila
	Version: 1.2.0
	#>
	[CmdletBinding(DefaultParameterSetName = 'Name')]
	param(
		[Parameter(Position = 0,
			ParameterSetName = 'Name',
			Mandatory = $true,
			ValueFromPipeline)]
		[ValidateNotNullOrEmpty()]
		[string[]]$Name,
		
		[Parameter(Position = 0,
			ParameterSetName = 'inPath',
			Mandatory = $true,
			ValueFromPipeline)]
		[ValidateScript( {
				if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
					Write-Host "File does not exist"
					throw
				}#endif
				return $True
			} ) ]
		[System.IO.FileInfo]$inPath,
	
		[Parameter(Mandatory = $False, ParameterSetName = 'Help')]
		[switch]$Help = $False
	)
	
	BEGIN {
		#Skripta tehniskie mainīgie
		$CurVersion = "1.2.0"
		#Skritpa konfigurācijas datnes
		$__ScriptName	= $MyInvocation.MyCommand
	
		$__ScriptPath	= Split-Path (
			Get-Variable MyInvocation -Scope Script
		).Value.Mycommand.Definition -Parent
	
		$Output = @()
		$LogOutput = @()
		
		if ($Help) {
			Write-Host "`nVersion:[$CurVersion]`n"
			$text = Get-Command -Name "$__ScriptPath\$__ScriptName" -Syntax
			$text | ForEach-Object { Write-Host $($_) }
			Write-Host "For more info write <Get-Help $__ScriptName -Examples>"
			Exit
		}#endif
	
		Function Stop-Watch {
			[CmdletBinding()] 
			param (
				[Parameter(Mandatory = $True)]
				[ValidateNotNullOrEmpty()]
				[object]$Timer,
				[Parameter(Mandatory = $True)]
				[ValidateNotNullOrEmpty()]
				[string]$Name
			)
			$Timer.Stop()
			if ( $Timer.Elapsed.Minutes -le 9 -and $Timer.Elapsed.Minutes -gt 0 )
			{ $bMin = "0$($Timer.Elapsed.Minutes)" } else 
			{ $bMin = "$($Timer.Elapsed.Minutes)" }
	
			if ( $Timer.Elapsed.Seconds -le 9 -and $Timer.Elapsed.Seconds -gt 0 ) 
			{ $bSec = "0$($Timer.Elapsed.Seconds)" } else 
			{ $bSec = "$($Timer.Elapsed.Seconds)" }
	
			Write-Host "`r[TestOnline] done in $(
				if ( [int]$Timer.Elapsed.Hours -gt 0 ) {"$($Timer.Elapsed.Hours)`:$bMin hrs"}
				elseif ( [int]$Timer.Elapsed.Minutes -gt 0 ) {"$($Timer.Elapsed.Minutes)`:$bSec min"}
				else { "$($Timer.Elapsed.Seconds)`.$($Timer.Elapsed.Milliseconds) sec" }
				)" -ForegroundColor Yellow -BackgroundColor Black
		}
	
		function Test-OnlineFast {
			param(
				[Parameter(Mandatory, ValueFromPipeline)]
				[string[]]$ComputersName,
				$TimeoutMillisec = 2000
			)
	
			BEGIN {
	
				[Collections.ArrayList]$bucket = @()
		
				$StatusCode_ReturnValue = @{
					0     = 'Success'
					11001 = 'Buffer Too Small'
					11002 = 'Destination Net Unreachable'
					11003 = 'Destination Host Unreachable'
					11004 = 'Destination Protocol Unreachable'
					11005 = 'Destination Port Unreachable'
					11006 = 'No Resources'
					11007 = 'Bad Option'
					11008 = 'Hardware Error'
					11009 = 'Packet Too Big'
					11010 = 'Request Timed Out'
					11011 = 'Bad Request'
					11012 = 'Bad Route'
					11013 = 'TimeToLive Expired Transit'
					11014 = 'TimeToLive Expired Reassembly'
					11015 = 'Parameter Problem'
					11016 = 'Source Quench'
					11017 = 'Option Too Big'
					11018 = 'Bad Destination'
					11032 = 'Negotiating IPSEC'
					11050 = 'General Failure'
				}
	
				$statusFriendlyText = @{
					Name       = 'Status'
					Expression = { 
						if ( $null -eq $_.StatusCode ) {
							"Unknown"
						}
						else {
							$StatusCode_ReturnValue[([int]$_.StatusCode)]
						}
					}
				}
	
				$IsOnline = @{
					Name       = 'Online'
					Expression = { $_.StatusCode -eq 0 }
				}
	
				$DNSName = @{
					Name       = 'DNSName'
					Expression = { if ($_.StatusCode -eq 0) { 
							if ($_.Address -like '*.*.*.*') 
							{ [Net.DNS]::GetHostByAddress($_.Address).HostName } 
							else  
							{ [Net.DNS]::GetHostByName($_.Address).HostName } 
						}
					}
				}
			}
		
			PROCESS {
	
				$ComputersName | ForEach-Object {
					$null = $bucket.Add($_)
				}
			}
	
			END {
	
				$query = $bucket -join "' or Address='"
				
				Get-CimInstance -ClassName Win32_PingStatus -Filter "(Address='$query') and timeout=$TimeoutMillisec" |
				Select-Object -Property $DNSName, Address, $IsOnline, $statusFriendlyText, StatusCode
			}
		}
	
		#region Ielasām mainīgos
	
		if ( $PSCmdlet.ParameterSetName -eq 'Name' ) {
			$Computers = $Name | Get-Unique
		}
		else {
			$Computers = Get-Content -Path $InPath | Where-Object { $_ -ne "" } `
			| Where-Object { -not $_.StartsWith('#') }  | Sort-Object | Get-Unique
		}
	
		#endregion
	
		Write-Host "[TestOnline] got for testing [$($Computers.Count)] $(
			if($Computers.Count -eq 1){"computer"}else{"computers"}
		)"
	
		Write-Host -NoNewLine "Running jobs : " -ForegroundColor Yellow -BackgroundColor Black
		$jobWatch = [System.Diagnostics.Stopwatch]::startNew()
	}
	
	PROCESS {
		
		$Output = @(
	
			$Computers | ForEach-Object {
	
				Write-Host -NoNewline "." -ForegroundColor Yellow -BackgroundColor Black
				$GetTest = Test-OnlineFast -ComputersName $_
	
				if ( $GetTest.Status -eq "Success" ) {
	
					if ( Test-WSMan -ComputerName $_ -ErrorAction SilentlyContinue) {
	
						$parameter = @{
							ClassName           = 'Win32_NetworkAdapterConfiguration'
							ComputerName        = "$_"
							OperationTimeoutSec = 5
							ErrorAction         = 'Stop'
						}
		
						$GetAddress = Get-CimInstance  @parameter -Filter "IPEnabled='True'" |
						Where-Object -Property DefaultIPGateway -ne $null | 
						Select-Object IPAddress, MACAddress
		
						New-Object -TypeName psobject -Property @{
							PipedName    = [string]$GetTest.Address.ToLower()
							DNSName      = [string]$GetTest.DNSName.ToLower()
							AddDate      = [System.DateTime](Get-Date)
							MacAddress   = [string]$GetAddress.MacAddress
							IPAddress    = [string]$GetAddress.IPAddress[0]
							WinRMservice = $true
						}
					}
					else {
						$LogOutput += @("[TestOnline] WinRM service is not accessible on host [$Computers]")
					}
				}
				else {
					$LogOutput += @("[TestOnline] host [$Computers] is not online")
				}
			}
		)
	}
	
	END {
		
		Stop-Watch -Timer $jobWatch -Name JOBers
	
		$LogOutput | ForEach-Object { Write-Host "$_" }
	
		if ( $Output.count -gt 0) {
			# $Output | ConvertTo-Json | Out-File "get-compTestOnline.json" -Force
			$Output
		}
		else {
			$false
		}
	}
}
