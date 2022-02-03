<#
.SYNOPSIS
Skripts atvieglo administratora windows jaunināšanas procesu, tehnisko parametru apkopošanu, attālināto darbstaciju startēšanu, pārsāknēšanu un izslēgšanu.

.DESCRIPTION
Skripts nodrošina:
[*] Windows jauninājumu pārbaudi, uzstādīšanu un datortehnikas pārsāknēšanu, ja to pieprasa jauninājums,
[*] pārbaudi datortehnikas nepieciešamībai uz pārsāknēšanu,
[*] attālināto datortehnikas sāknēšanu, restartu un izslēgšanu,
[*] datortehnikas tehnikos parametru un uzstādītās programmatūras pārskata izveidi.

.PARAMETER Name
Obligāts lauks.
Norādam datora vārdu.

.PARAMETER InPath
Obligāts lauks.
Norādam datoru saraksta datnes atrašanās vietu.

.PARAMETER Check
Kopā ar [-Name] vai [-InPath].
Veic konkrētā vai sarakstā norādīto datoru pārbaudi uz windows jauninājumiem. Rezultātu izvada ekrānā

.PARAMETER Update
Kopā ar [-Name] vai [-InPath].
Windows jauninājumu uzstādīšana.

.PARAMETER AutoReboot
Tikai kopā ar [-Update].
Automātisks datortehnikas restarts, ja jauninājums to pieprasa.

.PARAMETER WakeOnLan
Tikai kopā ar [-Name].
Veic norādītās datortehnikas pamodināšanu ar Magic paketes palīdzību.

.PARAMETER Stop
Tikai kopā ar [-Name].
Attālināti apstādina (shutdown) norādīto datoru.

.PARAMETER Reboot
Tikai kopā ar [-Name].
Attālināti restartē norādīto datoru un gaida, kamēr dators būs gatavs Powershell komandu izpildei

.PARAMETER NoWait
Tikai kopā ar [-Reboot].
Negaida, kamēr dators veiks pārsāknēšanas procedūru.

.PARAMETER Trace
Tikai kopā ar [-Name].
Tiešaistē seko līdzi Windows update žurnalēšanas datnes satura izmaiņām.

.PARAMETER EventLog
Tikai kopā ar [-Name] vai [-InPath]
Veic sarakstā norādīto serveru pārbaudi uz windows jauninājumu notikumiem datoru sistēmas notikumu žurnālā.

.PARAMETER Days
Tikai kopā ar [-EventLog]
Norādam par kādu periodu pagātnē tiek skatīti notikumi. Pēc noklusējuma - 30 dienas.

.PARAMETER OutPath
Tikai kopā ar [-InPath] un [-EventLog]
Norāda datnes vārdu, kurā tiks ierakstīts skripta rezultāts. Ja parametrs nav norādīts, rezultāts tiek izvadīts uz ekrāna.

.PARAMETER Asset
Tikai kopā ar [-Name].
Sagatavo datora tehnisko parametru un uzstādītās programmatūras un to versiju pārskatu.

.PARAMETER Include
Tikai kopā ar [-Asset]
Atlasa programmatūru pēc norādītā paterna. Atbalsta wildcard parametrus - *,?,[a-z] un [abc].
Vairāk informācijas šeit: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-5.1

.PARAMETER Exclude
Tikai kopā ar [-Asset]
Atlasa programmatūru, izņemot norādītajam paternam atbilstošo. Atbalsta wildcard parametrus - *,?,[a-z] un [abc].
Vairāk informācijas šeit: https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-5.1

.PARAMETER Hardware
Tikai kopā ar [-Asset]
Pārskatā iekļauj datortehnikas tehniskos parametrus.

.PARAMETER NoSoftware
Tikai kopā ar [-Asset]
Pārskatā neiekļauj uzstādīto programmatūru.

.PARAMETER Install
Tikai kopā ar [-Name] vai [-InPath]
Norādam programmatūras pakotnes atrašanās vietu un tā tiek uzstādīta uz uz datortehnikas. Uzstādīšanas žurnalēšanas datne atrodama C:/temp mapē.

.PARAMETER Uninstall
Tikai kopā ar [-Name] vai [-InPath]
Norādam programmatūras unikālo identifikatoru (Identifying Number) un norādītā programmatūra tiek novēkta no datora.
Identifying Number atrodams [-Asset] programmatūras pārskatā.

.PARAMETER ScriptUpdate
Pārbauda vai nav jaunākas skripta versijas datne norādītajā skripta etalona mapē.
Ja atrod - kopē uz darba direktoriju un beidz darbu.

.PARAMETER Help
Izvada skripta versiju, iespējamās komandas sintaksi un beidz darbu.

.EXAMPLE
WindowUpdate.ps1 -Name EX00001
Pārbauda datoram EX00001 pieejamos windows jauninājumus

.EXAMPLE
WindowUpdate.ps1 EX00001 -Asset
Sagatavo un parāda ekrānā datora EX00001 tehniskos parametrus

.EXAMPLE
WindowUpdate.ps1 -InPath .\computers.txt -Asset
Sagatavo tehnisko parametru atskaiti Excel formātā visām .\computers.txt norādītajām vienībām

.EXAMPLE
WindowUpdate.ps1 -InPath .\computers.txt
Pārbauda .\computers.txt datnē norādītajai datortehnikai pieejamos windows jauninājumus

.EXAMPLE
WindowUpdate.ps1 -InPath .\servers.txt -IsPendingReboot
Pārbauda vai .\servers.txt datnē norādītajai datortehnikai ir nepieciešams restarts.
Skripta log failā ir norāde uz CSV datnes atrašanās vietu.

.EXAMPLE
WindowUpdate.ps1 -Name EX00001 -Update
Izveido darbstacijā ScheduledTask Windows update uzdevumu, kas izpildās nekavējoši un uzsāk datortehnikai pieejamo jauninājumu lejupielādi, uzstādīšanu.
Tiek ignorēts jauninājuma pieprasījums pēc datoretehnikas pārsāknēšanas. Nepieciešams patstāvīgi pārliecināties, ka jauninājums uzstādījies pilnā apjomā un pārsāknēt darbstaciju.

.EXAMPLE
WindowUpdate.ps1 -InPath .\servers.txt -Update -AutoReboot
Sarakstā norādītajiem serveriem tiek izveidots ScheduledTask Windows update uzdevums, kas izpildās nekavējoši un uzsāk datortehnikai pieejamo jauninājumu lejupielādi, uzstādīšanu.
Ja jauninājuma pilnīgai uzstādīšanai ir nepeiciešams restarts - serveri tiek restartēti pēc jauninājuma pieprasījuma.

.EXAMPLE
WindowUpdate.ps1 -Name reja -RemoteReboot
Attālināti tiek pārsāknēta dators reja. Pārsāknēšana nav atceļama un tiek izpildīta nekavējoši.

.EXAMPLE
WindowUpdate.ps1 -EventLog EX00001
Pārbauda datora EX00001 notikumu žurnālu. Rāda tikai kopsavilkumu.

.EXAMPLE
WindowUpdate.ps1 -EventLog .\computers.txt -Details
Sagatavo .\computers.txt datnē norādītajiem datoru notikumu žurnāla ierakstus. Parāda detalizētu atskaiti.
Norādītai datnei ir jāeksistē, pretējā gadījumā tiek izvadīta kļūda.

.EXAMPLE
WindowUpdate.ps1 -EventLog EX00001 -Days 7
Pārbauda datora EX00001 notikumu žurnālu par notikumiem pēdējās 7 dienās

.EXAMPLE
WindowUpdate.ps1 -Install "C:\install\notepad\npp.8.1.9.3.Installer.x64.exe" EX00001
Uzstāda norādīto programmatūras pakotni uz datora.

.EXAMPLE
WindowUpdate.ps1 -Uninstall '{F914A43C-9614-4100-B94E-BE2D5EC2E5E2}' EX00001
Noņem norādīto programmatūras pakotni no datora. Programmatūras unikālais identifikators atrodams Asste pārskatā.

.EXAMPLE
WindowUpdate.ps1 -ScriptUpdate
Skripta piespiedu pārbaude uz skripta jauninājumu pieejamību skripta etalona mapē, kas tiek norādīta skripta mainīgajā $UpdateDir.

.NOTES
	Author:	Viesturs Skila
	Version: 3.0.1
#>
[CmdletBinding(DefaultParameterSetName = 'InPathCheck')]
param (
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathCheck',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathUpdate',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathEventLog',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPathAsset',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPath4Install',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'InPath4Uninstall',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'WakeOnLanInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'RebootInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'StopInPath',
		HelpMessage = "Path of txt file with list of computers.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".txt") {
				Write-Host "The file specified in the path argument must be text file"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$InPath,

	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'Name4Uninstall',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'Name4Install',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameAsset',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameEventLog',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameTrace',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'RebootName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'StopName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'WakeOnLanName',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameUpdate',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 0, Mandatory = $True,
		ParameterSetName = 'NameCheck',
		HelpMessage = "Name of computer")]
	[ValidateNotNullOrEmpty()]
	[ValidateScript( {
			if ( [String]::IsNullOrWhiteSpace($_) ) {
				Write-Host "Enter name of computer"
				throw
			}#endif
			return $True
		} ) ]
	[string]$Name,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathCheck')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameCheck')]
	[switch]$Check = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathUpdate')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameUpdate')]
	[switch]$Update = $False,
	
	[Parameter(Position = 2, Mandatory = $False, ParameterSetName = 'InPathUpdate')]
	[Parameter(Position = 2, Mandatory = $False, ParameterSetName = 'NameUpdate')]
	[switch]$AutoReboot = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'WakeOnLanInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'WakeOnLanName')]
	[switch]$WakeOnLan = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'RebootInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$Reboot = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$NoWait = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'StopInPath')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'StopName')]
	[switch]$Stop = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'StopInPath')]
	[Parameter(Mandatory = $False, ParameterSetName = 'RebootInPath')]
	[Parameter(Mandatory = $False, ParameterSetName = 'StopName')]
	[Parameter(Mandatory = $False, ParameterSetName = 'RebootName')]
	[switch]$Force = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameTrace')]
	[switch]$Trace = $False,

	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameEventLog')]
	[switch]$EventLog = $False,
	
	[Parameter(Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameEventLog')]
	[int]$Days,
	
	[Parameter(Mandatory = $False, ParameterSetName = 'InPathEventLog')]
	[switch]$OutPath = $False,
	
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Position = 1, Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$Asset = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[SupportsWildcards()]
	[string]$Include = '*',

	[Parameter(Mandatory = $False, ParameterSetName = 'InPathAsset')]
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[SupportsWildcards()]
	[string]$Exclude = '',
	
	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$Hardware = $False,

	[Parameter(Mandatory = $False, ParameterSetName = 'NameAsset')]
	[switch]$NoSoftware = $False,

	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'InPath4Install',
		HelpMessage = "Path of exe or msi file.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".msi|.exe") {
				Write-Host "The file specified in the path argument must be msi or exe file"
				throw
			}#endif
			return $True
		} ) ]
	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'Name4Install',
		HelpMessage = "Path of exe or msi file.")]
	[ValidateScript( {
			if ( -NOT ( $_ | Test-Path -PathType Leaf) ) {
				Write-Host "File does not exist"
				throw
			}#endif
			if ( $_ -notmatch ".msi|.exe") {
				Write-Host "The file specified in the path argument must be msi or exe file"
				throw
			}#endif
			return $True
		} ) ]
	[System.IO.FileInfo]$Install,

	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'InPath4Uninstall')]
	[Parameter(Position = 1, Mandatory = $True, ParameterSetName = 'Name4Uninstall')]
	[string]$Uninstall,

	#Helper slēdži
	[Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'ScriptUpdate')]
	[switch]$ScriptUpdate = $False,

	[Parameter(Position = 0, Mandatory = $False, ParameterSetName = 'Help')]
	[switch]$Help = $False
)
BEGIN {
	# $UpdateDir = "\\beluga\install\Scripts\ExpoRemoteJobs"
	#$PsBoundParameters
	<# ---------------------------------------------------------------------------------------------------------
		Zemāk veicam izmaiņas, ja patiešām saprotam, ko darām.
	--------------------------------------------------------------------------------------------------------- #>
	#Skripta tehniskie mainīgie
	$CurVersion = "3.0.1"
	$scriptWatch	= [System.Diagnostics.Stopwatch]::startNew()
	#Skritpa konfigurācijas datnes
	$__ScriptName	= $MyInvocation.MyCommand
	$__ScriptPath	= Split-Path (Get-Variable MyInvocation -Scope Script).Value.Mycommand.Definition -Parent
	#Atskaitēm un datu uzkrāšanai
	$ReportPath = "$__ScriptPath\result"
	$Script:DataDir = "$__ScriptPath\data"
	$BackupDir = "$Script:DataDir\backup"
	# $ModuleDir = "$__ScriptPath\modules"
	$LogFileDir = "$__ScriptPath\log"
	# ielādējam moduļus
	Get-PSSession | Remove-PSSession
	Remove-Module -Name VSkWinUpdate -ErrorAction SilentlyContinue
	Import-Module -Name ".\modules\VSkWinUpdate.psm1"
	#Helper scriptu bibliotēkas
	# $CompUpdateFileName = "Set-CompUpdate.ps1"
	# $CompProgramFileName = "Set-CompProgram.ps1"
	# $CompAssetFileName = "Get-CompAsset.ps1"
	# $CompSoftwareFileName	= "Get-CompSoftware.ps1"
	# $CompEventsFileName = "Get-CompEvents.ps1"
	# $CompTestOnlineFileName	= "Get-CompTestOnline.ps1"
	# $CompWakeOnLanFileName	= "Invoke-CompWakeOnLan.ps1"
	# $WinUpdFile = "$ModuleDir\$CompUpdateFileName"
	# $CompProgramFile = "$ModuleDir\$CompProgramFileName"
	# $CompAssetFile	= "$ModuleDir\$CompAssetFileName"
	# $CompSoftwareFile = "$ModuleDir\$CompSoftwareFileName"
	# $CompEventsFile = "$ModuleDir\$CompEventsFileName"
	# $CompTestOnlineFile	= "$ModuleDir\$CompTestOnlineFileName"
	# $CompWakeOnLanFile	= "$ModuleDir\$CompWakeOnLanFileName"
	#Žurnalēšanai
	$PSDefaultParameterValues['out-file:width'] = 500
	$Script:LogFile = "$LogFileDir\$__ScriptName-$(Get-Date -Format "yyyyMMdd")"
	#$XLStoFile		= "$ReportPath\Report-$(Get-Date -Format "yyyyMMddHHmmss").xls"
	$OutputToFile	= "$ReportPath\$__ScriptName-$(Get-Date -Format "yyyyMMdd").txt"
	$TraceFile = "C:\ExpoSheduledWUjob.log"
	$DataArchiveFile = "$($Script:DataDir)\DataArchive.dat"
	#$ComputerName = ( -not [string]::IsNullOrEmpty($Name) )
	$ScriptUser = Invoke-Command -ScriptBlock { whoami }
	$RemoteComputers = @()
	$Script:MaxJobsThreads	= 40


	if ( -not ( Test-Path -Path $LogFileDir ) ) { $null = New-Item -ItemType "Directory" -Path $LogFileDir }
	if ( -not ( Test-Path -Path $ReportPath ) ) { $null = New-Item -ItemType "Directory" -Path $ReportPath }
	if ( -not ( Test-Path -Path $Script:DataDir ) ) { $null = New-Item -ItemType "Directory" -Path $Script:DataDir }
	if ( -not ( Test-Path -Path $BackupDir ) ) { $null = New-Item -ItemType "Directory" -Path $BackupDir }
	

	if ($Help) {
		Get-ScriptHelp -Version $CurVersion -ScriptPath "$__ScriptPath\$__ScriptName"
		Exit
	}#endif

	<# ---------------------------------------------------------------------------------------------------------
	Funkciju definēšanu beidzām
	IELASĀM SCRIPTA DARBĪBAI NEPIECIEŠAMOS PARAMETRUS
	--------------------------------------------------------------------------------------------------------- #>
	Write-msg -log -text "[-----] Script started in [$(if ($Trace) {"Trace"}
		elseif ($ScriptUpdate) {"ScriptUpdate"}
		elseif ($Asset) {"Asset"}
		elseif ($Update) {"Update"} 
		elseif ($Reboot) {"Reboot"}
		elseif ($Stop) {"Stop"} 
		elseif ( $PSCmdlet.ParameterSetName -eq "NameEventLog" ) {"Name-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -eq "InPathEventLog" ) {"InPath-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLan*" ) {"WakeOnLan"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Install" ) {"Install"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Uninstall" ) {"Uninstall"}
		else {"Check"})] mode. Used value [$(if ($Name){"Name:[$Name]"}
			elseif ($InPath) {"InPath:[$InPath]"} 
			else {"none"})]"
	Write-Host "`n[-----] Script started in [$(if ($Trace) {"Trace"}
		elseif ($ScriptUpdate) {"ScriptUpdate"}
		elseif ($Asset) {"Asset"}
		elseif ($Update) {"Update"} 
		elseif ($Reboot) {"Reboot"}
		elseif ($Stop) {"Stop"}
		elseif ( $PSCmdlet.ParameterSetName -eq "NameEventLog" ) {"Name-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -eq "InPathEventLog" ) {"InPath-EventLog"}
		elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLan*" ) {"WakeOnLan"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Install" ) {"Install"}
		elseif ( $PSCmdlet.ParameterSetName -like "*4Uninstall" ) {"Uninstall"}
		else {"Check"})] mode. Used value [$(if ($Name){"Name:[$Name]"}
			elseif ($InPath) {"InPath:[$InPath]"} 
			else {"none"})]"
			
	#Scripts pārbauda vai nav jaunākas versijas repozitorijā 
	# if ( Test-Path -Path $UpdateDir ) {
	# 	Get-ScriptFileUpdate $__ScriptName $__ScriptPath $UpdateDir
	# 	Get-ScriptFileUpdate $CompUpdateFileName $__ScriptPath $UpdateDir
	# 	Get-ScriptFileUpdate $CompProgramFileName $__ScriptPath $UpdateDir
	# 	# Get-ScriptFileUpdate $CompAssetFileName $__ScriptPath $UpdateDir
	# 	# Get-ScriptFileUpdate $CompSoftwareFileName $__ScriptPath $UpdateDir
	# 	# Get-ScriptFileUpdate $CompEventsFileName $__ScriptPath $UpdateDir
	# 	# Get-ScriptFileUpdate $CompTestOnlineFileName $__ScriptPath $UpdateDir
	# 	# Get-ScriptFileUpdate $CompWakeOnLanFileName $__ScriptPath $UpdateDir
	# 	if ($ScriptUpdate) {
	# 		Stop-Watch -Timer $scriptWatch -Name Script
	# 		exit 
	# 	}#endif
	# }#endif
	# else {
	# 	if ($ScriptUpdate) {
	# 		Write-msg -log -text "[Update] directory [$($UpdateDir)] not available."
	# 		Stop-Watch -Timer $scriptWatch -Name Script
	# 		exit 
	# 	}#endif
	# }#endif
	
	# #Check for script helper files; if not aviable - exit
	# if ( -not ( Test-Path -Path $WinUpdFile -PathType Leaf )) {
	# 	Write-msg -log -bug -text "File [$WinUpdFile] not found. Exit."
	# 	Stop-Watch -Timer $scriptWatch -Name Script
	# 	Exit
	# }#endif
	# if ( -not ( Test-Path -Path $CompTestOnlineFile -PathType Leaf )) {
	# 	Write-msg -log -bug -text "File [$CompTestOnlineFile] not found. Exit."
	# 	Stop-Watch -Timer $scriptWatch -Name Script
	# 	Exit
	# }#endif
	
	<# ---------------------------------------------------------------------------------------------------------
	pārbaudām katras datortehnikas gatavību strādāt PSRemote režīmā, vai ir nepieciešamās bilbiotēkas
	To darām trīs soļos: 
	[1] ielādējam sarakstu un pingojam,
	[2] ielādējam arhīvu un pārbaudam vai TTL nav beidzies,
	[3] ja TTL beidzies, veicam datortehnikas pilno pārbaudi uz PSRemote
	--------------------------------------------------------------------------------------------------------- #>
	
	#region ielādējam input parametrus
	$paramCompTestOnline = @{}
	if ( $PSCmdlet.ParameterSetName -like "Name*" -or 
		$PSCmdlet.ParameterSetName -eq "RebootName" -or 
		$PSCmdlet.ParameterSetName -eq "StopName" -or 
		$PSCmdlet.ParameterSetName -like "Asset*" ) {
		$InComputers = @($Name.ToLower())
		# $ArgumentListOnline = "`-Name $InComputers"
		$paramCompTestOnline.Add('Name', $InComputers)
	}#endif
	elseif ( $PSCmdlet.ParameterSetName -like "InPath*" -or
		$PSCmdlet.ParameterSetName -eq "RebootInPath" -or 
		$PSCmdlet.ParameterSetName -eq "StopInPath" ) {
		
		
		$InComputers = @(Get-Content $InPath | 
			Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  |
			ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique )

		# $ArgumentListOnline = "`-Inpath $InPath"
		$paramCompTestOnline.Add('Inpath', $InPath)
	}#endelseif
	elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLanName" ) {
		$RemoteComputers = @($Name.ToLower())
	}
	elseif ( $PSCmdlet.ParameterSetName -like "WakeOnLanInPath" ) {
		$RemoteComputers = @(Get-Content $InPath | Where-Object { $_ -ne "" } | Where-Object { -not $_.StartsWith('#') }  | ForEach-Object { $_.ToLower() } | Sort-Object | Get-Unique )
	}

	#endregion
	
	if ( $PSCmdlet.ParameterSetName -like "Name*" -or 
		$PSCmdlet.ParameterSetName -like "InPath*" -or
		$PSCmdlet.ParameterSetName -like "Reboot*" -or 
		$PSCmdlet.ParameterSetName -like "Stop*" -or 
		$PSCmdlet.ParameterSetName -like "Asset*" ) {
		#ielādējam datu arhīvu
		if ( Test-Path $DataArchiveFile -PathType Leaf ) {
			try {
				$DataArchive = @(Import-Clixml -Path $DataArchiveFile -ErrorAction Stop)
			}#endtry
			catch {
				$DataArchive = @()
			}#endcatch
		}#endif
		else {
			$DataArchive = @()
		}#endif
		
		Write-Verbose "1.0:[Input]:got:[$($InComputers.count)]"
		<#
		Write-Verbose "1.1:[DataArchive]-=> []---------------------------------------------------------"
		$DataArchive | Sort-Object -Property AddDate -Descending `
		| Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize  `
		| Out-String -Stream | Where-Object { $_ -ne "" } `
		| ForEach-Object { Write-Verbose "$_" }
		Write-Verbose "--------------------------------------------------------------------------------"
		#>

		$OfflineComputers = @()
		Write-Verbose "`& `"$CompTestOnlineFile`" $ArgumentListOnline "
		
		# Izsaucam Get-CompTestOnline.ps1
		$OnlineComps = Get-CompTestOnline @paramCompTestOnline

		Write-Verbose "OnlineComps:[$(if ( $OnlineComps ) {"Atgriezts masīvs"}  else {"Atgriezts tukšs masīvs"})]"
		if ( $OnlineComps ) {
			Write-Verbose "1.2:[OnlineComps]---------------------------------------------------------------"
			$OnlineComps | Sort-Object -Property AddDate -Descending |
			Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize | 
			Out-String -Stream | Where-Object { $_ -ne "" } |
			ForEach-Object { Write-Verbose "$_" }
			Write-Verbose "--------------------------------------------------------------------------------"
			Write-Verbose "1.3:OnlineComps:[$(if ( $OnlineComps.GetType().BaseType.name -eq 'Array' -and $OnlineComps.count -gt 0 ) {"Array"} `
				elseif ( $OnlineComps.GetType().BaseType.name -eq 'Object' ) {"Object"} else {"Other"})]; DataArchive:[$(`
				if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -gt 0 ) {"Array"} `
				elseif ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -eq 0 ) {"Empty Array"} `
				elseif ( $DataArchive.GetType().BaseType.name -eq 'Object' ) {"Object"} else {"Other"})]`
				"
		}
		
		#Ja OnlineComps ir objekts, tad to pārveidojam to par objektu masīvu
		if ( $OnlineComps.GetType().BaseType.name -eq 'Object' ) {
			$tmpOnlineComps = $OnlineComps.psobject.copy()
			$OnlineComps = @()
			$OnlineComps += @($tmpOnlineComps)
		}

		if ( $OnlineComps -eq $false) {
			Write-Verbose "[Get-CompTestOnline] returned an empty object"
			$RemoteComputers = @()
		}
		#ja atgriezts objektu masīvs
		elseif ( $OnlineComps.GetType().BaseType.name -eq 'Array' -and $OnlineComps.count -gt 0 ) {
			#Atlasām offline datorus
			$InComputers | ForEach-Object {
				if ( $OnlineComps.PipedName.Contains($_) -eq $false) {
					Write-Verbose "1:[OnlineComps]:[$($_)] -=> [OfflineComputers]] "
					$OfflineComputers += @($_)
				}
			}
			[string]$vstring = $null
			$OnlineComps.DNSName | ForEach-Object { $vstring += "[$_], " }
			Write-Verbose "2.0:[OnlineComps]: $vstring"
			Write-Verbose "OfflineComputers:[$OfflineComputers]"
			
			$_OnlineComps = @()
			<# ---------------------------------------------------------------------------------------------------------
				pārbaudam ARHĪVĀ vai datoram  nav beidzies TTL derīgums: 12 stundas)
			--------------------------------------------------------------------------------------------------------- #>
			#izlaižam arhīva soli, ja [1] arhīvs nesatur vērtības vai [2] Online ir tikai viens ieraksts
			if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and
				$DataArchive.count -gt 0 -and
				$OnlineComps.count -gt 1 ) {

				foreach ( $comp in $OnlineComps) {
					:OutOfNestedForEach_LABEL
					foreach ( $record in $DataArchive) {
						if ( $comp.DNSName -eq $record.DNSName ) {
							#ja dators pēdējo reizi verificēts pirms 12h, tad pārbaudām
							if ($record.AddDate -lt (Get-Date).AddHours(-12) ) {
								$_OnlineComps += @($record.DNSName)
								Write-Verbose "2.1:[Archive]:[$($record.DNSName)]-=> [_OnlineComps]"
								break :OutOfNestedForEach_LABEL
							}
							else {
								$RemoteComputers += @($record.DNSName)
								Write-Verbose "2.2:[Archive]:[$($record.DNSName)]-=> [RemoteComputers]"
								break :OutOfNestedForEach_LABEL
							}
						}
					}
				}

				#pārbaudam uz online ierakstu esamību, kas netika atrasti arhīvā - ja ir, tad liekam _Online
				$OnlineComps.DNSName | ForEach-Object {
					if ( $_OnlineComps.Contains($_) -eq $false -and $RemoteComputers.Contains($_) -eq $false ) {
						Write-Verbose "2.3:[Online]:[$($record.DNSName)]-=> [RemoteComputers]"
						$_OnlineComps += @($_)
					}
				}
			}
			else {
				$_OnlineComps = $OnlineComps | ForEach-Object { @($_.DNSName) }
			}
			Write-Verbose "2.4:_OnlineComps:[$_OnlineComps]; RemoteComputers[$RemoteComputers]"

			<# ---------------------------------------------------------------------------------------------------------
				ja atrasti ieraksti, kam TTL beidzies, veicam datortehnikas padziļināto atbilstības pārbaudi
			--------------------------------------------------------------------------------------------------------- #>
			if ( $_OnlineComps.count -gt 0 ) {

				# Pārbaudam datoru Powershell iestatījumu atbilstību 
				$VerifiedComps = Test-VSkRemoteComputer -ComputerName $_OnlineComps
				
				Write-Verbose "3.0: Got from VerifiedComps:[$VerifiedComps]"

				#Ja pārbaudi neiztur neviens dators
				if ( $VerifiedComps -eq $false ) {
					$OfflineComputers += $_OnlineComps | ForEach-Object { @($_) }
				}
				elseif ( $VerifiedComps.count -gt 0 ) {
					foreach ( $comp in $OnlineComps) { 
						Write-Verbose "3.1:[OnlineComps] [$($comp.DNSName)]:[$($comp.PipedName)]:[$($comp.AddDate)]:[$($comp.MacAddress)]"
						#Ja izturēja pārbaudi - liekam remote un papildinām arhīvu, ja ne - offline
						if ( $VerifiedComps.Contains($comp.DNSName) -eq $true ) {
							Write-Verbose "3.2:[OnlineComps]:[$($comp.DNSName)]-=> [RemoteComputers]"
							$RemoteComputers += @($comp.DNSName)
							#salāgojam online un arhīva ierakstus
							if ( $DataArchive.GetType().BaseType.name -eq 'Array' -and $DataArchive.count -eq 0 ) {
								Write-Verbose "3.3:[OnlineComps]:[$($comp.DNSName)]-=> [DataArchive]"
								$DataArchive = @(
									New-Object PSObject -Property @{
										PipedName    = $comp.PipedName.ToLower();
										DNSName      = $comp.DNSName.ToLower();
										AddDate      = [System.DateTime](Get-Date);
										MacAddress   = $comp.MacAddress;
										IPAddress    = $comp.IPAddress;
										WinRMservice = $comp.WinRMservice;
									}#endobject
								)
							}#endif
							elseif ( $DataArchive.DNSName.Contains($comp.DNSName) -eq $false ) {
								Write-Verbose "3.4:[OnlineComps]:[$($comp.DNSName)]-=> [DataArchive]"
								$DataArchive += @(
									New-Object PSObject -Property @{
										PipedName    = $comp.PipedName.ToLower();
										DNSName      = $comp.DNSName.ToLower();
										AddDate      = [System.DateTime](Get-Date);
										MacAddress   = $comp.MacAddress;
										IPAddress    = $comp.IPAddress;
										WinRMservice = $comp.WinRMservice;
									}#endobject
								)
							}#endelseif
							foreach ( $row in $DataArchive ) {
								if ( $comp.DNSName -eq $row.DNSName ) {
									$row.AddDate = [System.DateTime](Get-Date)
								}
							}#endforeach
						}#endif
						else {
							if ( $RemoteComputers.Contains($comp.DNSName) -eq $false ) {
								Write-Verbose "3.5:[OnlineComps]:[$($_)]-=> [OfflineComputers] "
								$OfflineComputers += @($comp.DNSName)
							}#endif
						}#endelse
					}#endforeach
				}#endelseif
				else {
					$OfflineComputers += $_OnlineComps | ForEach-Object { @($_) }
				}#endelse
			}#endif
			else {
				Write-Verbose "3.0:[VerifiedComps]: -= SKIPPED =-"
			}#endif
		}#endif
		else {
			Write-msg -log -bug -text "[OnlineComps] returned Object type [Other] or [Empty]"
			$RemoteComputers = @()
		}#endelse
		<#
		Write-Verbose "4.0:[]-=> [DataArchive] --------------------------------------------------------"
		$DataArchive | Sort-Object -Property AddDate -Descending `
		| Format-Table AddDate, PipedName, DNSName, MacAddress -AutoSize `
		| Out-String -Stream | Where-Object { $_ -ne "" } `
		| ForEach-Object { Write-Verbose "$_" }
		Write-Verbose "--------------------------------------------------------------------------------"
		#>
		#ierakstām datus arhīvā
		$parameter = @{
			Path        = $DataArchiveFile
			Destination = "$BackupDir\DataArchive-$(Get-Date -Format "yyyyMMddHHmm").bck"
			ErrorAction = 'SilentlyContinue'
		}

		Copy-Item @parameter

		$DataArchive | Export-Clixml -Path $DataArchiveFile -Depth 10 -Force

		Write-Verbose "[Online]:[$($RemoteComputers.count)],[Offline]:[$($OfflineComputers.Count)]"
	}#endif
}#endOfbegin

<# ---------------------------------------------------------------------------------------------------------
	ŠEIT SĀKAS PAMATA DARBS
--------------------------------------------------------------------------------------------------------- #>
PROCESS {
	<# ---------------------------------------------------------------------------------------------------------
		Reboot vai Stop - prasām apliecinājumus un darām darbu
	--------------------------------------------------------------------------------------------------------- #>
	#region REBOOT and STOP

	if ( ( $RemoteComputers.count -gt 0 ) -and `
		( $PSCmdlet.ParameterSetName -like "Reboot*" -or `
				$PSCmdlet.ParameterSetName -like "Stop*" ) ) {
		
		Write-Host "Please be sure in what you are going to do!!!`n---------------------------------------------" -ForegroundColor Yellow
		foreach ( $computer in $RemoteComputers ) {
			Write-msg -log -text "[Main] Confirm $(if($Reboot) {"reboot"} else {"shutdown"}) of [$computer]"
			$answer = Read-Host -Prompt "Please confirm remote $(if($Reboot) {"reboot"} else {"shutdown"}) of [ $computer ]:`t`tType [Yes/No]"
			if ( $answer -like 'Yes' ) {
				Write-msg -log -text "[Main] [$Computer] going to be $(if($Reboot) {"rebooted"} else {"shutdowned"})."
				if ($Reboot) {
					try {
						if ($NoWait -or $PSCmdlet.ParameterSetName -eq "RebootInPath") {
							$parameters = @{
								ComputerName = $Computer
							}
						}
						else {
							$parameters = @{
								ComputerName = $Computer
								Wait         = $True
								For          = 'Powershell'
								Timeout      = 300
								Delay        = 2 
							}
						}
						if ($Force) {
							$parameters.Add( 'Force', $Force ) 
						}
						$parameters.Add( 'ErrorAction', 'Stop' )

						Restart-Computer @parameters
						Write-msg -log -text "[Reboot] [$Computer] successfully."
						Write-Host "[Reboot] [$Computer] successfully."
					}
					catch {
						Write-ErrorMsg -Name 'Reboot' -InputObject $_
					}
				}
				if ($Stop) {
					try {
						$parameters = @{
							ComputerName = $Computer
							ErrorAction  = 'Stop'
						}#endsplat
						if ($Force) { 
							$parameters.Add( 'Force', $Force ) 
						}#endif

						Stop-Computer @parameters
						Write-msg -log -text "[Stop] [$Computer] successfully."
						Write-Host "[Stop] [$Computer] successfully."
					}
					catch {
						Write-ErrorMsg -Name 'Stop' -InputObject $_
					}
				}
			}
			else {
				Write-msg -log -text "[$(if($Reboot) {"Reboot"} else {"Stop"})] [$Computer] $(if($Reboot) {"reboot"} else {"shutdown"}) canceled."
				Write-Host "[$(if($Reboot) {"Reboot"} else {"Stop"})] [$Computer] $(if($Reboot) {"reboot"} else {"shutdown"}) canceled."
			}
		}
	}
	#endregion

	<# ---------------------------------------------------------------------------------------------------------
		izpildam Asset skriptu			- $CompAssetFileName	= "lib\Get-CompAsset.ps1"
		izpildam Get-Software skriptu	- $CompSoftwareFileName = "lib\Get-CompSoftware.ps1"
	--------------------------------------------------------------------------------------------------------- #>
	#region ASSET: SOFTWARE and HARDWARE
	if ( $RemoteComputers.count -gt 0 -and $Asset ) {
		try {
			$CompSession = New-PSSession -ComputerName $RemoteComputers -ErrorAction Stop
			Write-msg -log -text "[Asset]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
			Write-Host "[Asset]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"

			if ( $PSCmdlet.ParameterSetName -eq "NameAsset" ) {
				if ( $Hardware ) {
					
					Write-Host "`n[Computer]====================================================================================================" -ForegroundColor Yellow

					$result = Invoke-Command -Session $CompSession -ScriptBlock ${Function:Get-CompHardware} -ArgumentList $Hardware
					
					$result | Format-List | Out-String -Stream | Where-Object { $_ -ne "" } |
					ForEach-Object { Write-Verbose "$_" }

				}
				if ( -NOT $NoSoftware ) {
					Write-Host "`n[Software]====================================================================================================" -ForegroundColor Yellow

					$result = Invoke-Command -Session $CompSession -ScriptBlock ${Function:Get-CompSoftware} -ArgumentList ($Include, $Exclude) |
					Sort-Object -Property DisplayName |
					Select-Object PSComputerName, @{ name = 'Name'; expression = { $_.DisplayName } }, 
					@{name = 'Version'; expression = { $_.DisplayVersion } }, Scope, IdentifyingNumber, 
					@{name = 'Arch'; expression = { $_.Architecture } } -Unique

					if ($result) {
						$result	| Format-Table Name, Version, IdentifyingNumber, Scope, Arch -AutoSize
					}
					else {
						Write-Host "Got nothing to show." -ForegroundColor Green
					}

				}
				Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
			}
			if ($PSCmdlet.ParameterSetName -eq "InPathAsset" ) {

				Write-Host "`n[Software]====================================================================================================" -ForegroundColor Yellow
				
				$result = Invoke-Command -Session $CompSession -ScriptBlock ${Function:Get-CompSoftware} -ArgumentList ($Include, $Exclude) | 
				Sort-Object -Property PSComputerName, DisplayName |
				Select-Object PSComputerName, 
				@{name = 'Name'; expression = { $_.DisplayName } }, 
				@{name = 'Version'; expression = { $_.DisplayVersion } }, 
				Scope, IdentifyingNumber, 
				@{name = 'Arch'; expression = { $_.Architecture } } -Unique

				if ($result) {
					$result	| Format-Table PSComputerName, Name, Version, IdentifyingNumber, Scope, Arch -AutoSize
				}
				else {
					Write-Host "Got nothing to show." -ForegroundColor Green
				}
				Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
				<#
				if ( $RawData.count -gt 0 ) {
					Write-msg -log -text "[Asset] Collecting information..."
					$ReportToExcel = Get-NormaliseDiskLabelsForExcel $RawData
					$properties = $ReportToExcel | Foreach-Object { $_.psobject.Properties | Select-Object -ExpandProperty Name } | Sort-Object -Unique
					#for testing
					$ReportToExcel | Sort-Object -Property aComputerName | Select-Object $properties | `
						ConvertTo-Json -depth 10 | Out-File "$DataDir\Data-$(Get-Date -Format "yyyyMMddHHmmss").json" -Force
					$ReportToExcel | Sort-Object -Property aComputerName | Select-Object $properties | `
						Export-Excel $XLStoFile -WorksheetName "ExpoAssets" -FreezeTopRow -AutoSize -BoldTopRow
					Write-msg -log -text "Please check the Excel file [$XLStoFile]."
					Write-Host "[Asset] Please check the Excel file [$XLStoFile].`n "
				}#endif
				else {
					Write-msg -log -bug -text "[Asset] Ups... there's nothing to import to the Excel."
				}#endelse
				#>
			}
		}
		catch {
			Write-ErrorMsg -Name 'Asset' -InputObject $_
		}
		finally {
			if ( $CompSession.count -gt 0 ) {
				Remove-PSSession -Session $CompSession
			}
		}
	}
	else {
		if ( $PSCmdlet.ParameterSetName -like "*Asset" ) {
			Write-msg -log -bug -text "[Asset] No computer in list."
		}
	}

	#endregion

	<# ---------------------------------------------------------------------------------------------------------
		CHECK, UPDATE and TRACE
	--------------------------------------------------------------------------------------------------------- #>
	#region WINDOWS UPDATE: CHECK, UPDATE and TRACE

	if ( ( $RemoteComputers.count -gt 0 ) -and ( $Check -or $Update -or $Trace ) ) {
		try {
			Write-msg -log -text "Conecting to [$($RemoteComputers.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} ).."
			$CompSession = New-PSSession -ComputerName $RemoteComputers -ErrorAction Stop

			Write-msg -log -text "[$(if($Check){"Check"}elseif($Update){"Update"})]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
			Write-Host "[$(if($Check){"Check"}elseif($Update){"Update"})]:got:[$($RemoteComputers.count)] -=> [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"

			if ($CompSession.count -gt 0 ) {

				#region TRACE
				if ( $Trace) {
					
					$TestScript = {
						$TraceFile = $args[0]
						$isFile = Test-Path -Path $TraceFile -PathType Leaf
						return $isFile
					}
					$SessScript = {
						$TraceFile = $args[0]
						$file = Get-Content $TraceFile
						return $file
					}
					# check is there log file source; if not - skip
					if ( Invoke-Command -Session $CompSession -ScriptBlock $TestScript -ArgumentList $TraceFile ) {
						$result = Invoke-Command -Session $CompSession -ScriptBlock $SessScript -ArgumentList $TraceFile -ErrorAction Ignore
						Write-host "`n[$(Get-Date -Format "yyyy.MM.dd HH:mm:ss")] from [$TraceFile]--------------------------" -ForegroundColor Yellow
			
						$result | Format-Table * -AutoSize | Out-String -Width 128 -Stream `
						| Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" }
							
						Write-Host "-------------------------------------------------------------------------------" -ForegroundColor Yellow

						$result | ConvertTo-Json -depth 10 | Out-File "$DataDir\trace.json"
			
					}
					else {
						Write-Host "There's no update process started." -Foreground Yellow 
						Write-msg -log -text "There's no update process started."
					}

				}
				#endregion

				#region CHECK and UPDATE
				if ($Check -or $Update) {
					Write-msg -log -text "Sending instructions to [$($CompSession.count)] $(if ( $RemoteComputers.count -gt 1 ) {"computers"} else {"computer"} )"
					# Write-Host "[WindowsUpdate] Update:[$($Update)], AutoReboot[$($AutoReboot)]"
					# Run Windows install script on to each computer
					$paramVSkJob = @{
						Computers       = $RemoteComputers
						ScriptBlockName = 'Set-CompWindowsUpdate'
						Update          = $Update
						AutoReboot      = $AutoReboot
					}

					$JobResults = Send-VSkJob @paramVSkJob

					Write-Host "=================================================================================================" -ForegroundColor Yellow
					
					# Collect remote logs
					$ResultOutput = @()
					$i = 0 
					$JobResults = $JobResults | Sort-Object -Property Computer
					Foreach ( $row in $JobResults ) {
						#Write-Host "1: [$($row.PendingReboot.Count)][$(if ( $row.PendingReboot) {"True"})][$(if ( $null -eq $row.PendingReboot ) {"True"})]"
						if ( $null -ne $row.PendingReboot ) {
							$row.PendingReboot | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									#Examle of line: "EUROBARS | update | [4] updates are waiting to be installed"
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}
								$i++
							}
						}
						#Write-Host "2: [$($row.Updates.Count)][$(if ( $row.Updates) {"True"})][$(if ( $null -eq $row.Updates ) {"True"})]"
						if ( $null -ne $row.Updates ) {
							$row.Updates | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}
								$i++
							}
						}
						#Write-Host "3: [$($row.ScheduledTask.Count)][$(if ( $row.ScheduledTask) {"True"})][$(if ( $null -eq $row.ScheduledTask ) {"True"})]"
						if ( $null -ne $row.ScheduledTask ) {
							$row.ScheduledTask | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}
								$i++
							}
						}
						#Write-Host "4: [$($row.ErrorMsg.Count)][$(if ( $row.ErrorMsg) {"True"})][$(if ( $null -eq $row.ErrorMsg ) {"True"})]"
						if ( $null -ne $row.ErrorMsg ) {
							$row.ErrorMsg | ForEach-Object {
								$ResultOutput += New-Object -TypeName psobject -Property @{
									id      = $i
									Name    = (($_.Split('|'))[0]).trim(); ;
									Title   = (($_.Split('|'))[1]).trim();
									Message	= (($_.Split('|'))[2]).trim();
								}
								$i++
							}
						}
					}
					if ( $PSCmdlet.ParameterSetName -like "Name[cu]*" ) {
						Write-Host "`nReport:"
						Write-Host "======="
						$ResultOutput | Sort-Object -Property id | Format-Table -Property Name, Title, Message -AutoSize `
						| Out-String -Stream | Where-Object { $_ -ne "" } | ForEach-Object { Write-Host "$_" }
					}
					else {
						Write-Output "`n[$(Get-Date -Format "yyyy.MM.dd HH:mm:ss")][$ScriptUser][$InPath]-------------------------" `
						| Out-File -FilePath $OutputToFile -Encoding 'ASCII' -Append -Force	
						# OutputToFile
						$ResultOutput | Sort-Object -Property id | Format-Table -Property Name, Title, Message -AutoSize `
						| Out-String -Stream | Where-Object { $_ -ne "" } `
						| Out-File -FilePath $OutputToFile -Encoding 'ASCII' -Append -Force	
						Write-Host "The report file is [ $OutputToFile ]." 
						Write-msg -log -text "[Main] The report file is [ $OutputToFile ]."
					}

				}
			}

			#endregion

		}
		catch {
			Write-ErrorMsg -Name 'JObRunners' -InputObject $_
		}
		finally {
			if ( $CompSession.count -gt 0 ) {
				Remove-PSSession -Session $CompSession
			}
		}
	}
	else {
		if ( $PSCmdlet.ParameterSetName -like "Name[cu]*" -or $PSCmdlet.ParameterSetName -like "InPath[cu]*" ) {
			Write-msg -log -bug -text "[$(if($Check){"Check"}elseif($Update){"Update"})] No computer in list."
		}
	}

	#endregion

	<# ---------------------------------------------------------------------------------------------------------
		Izsaucam EventLog skriptu: NameEventLog or InPathEventLog; $CompEventsFileName 	= "lib\Get-CompEvents.ps1"
		# Get-CompEvents.ps1 [-InPath] <FileInfo> [-OutPath <switch>] [-Days <int>] [<CommonParameters>]
		# Get-CompEvents.ps1 [-ComputerName] <string> [-Named <switch>][-Days <int>] [<CommonParameters>]
	--------------------------------------------------------------------------------------------------------- #>
	#region  EVENTLOG
	if ( ( $RemoteComputers.count -gt 0 ) -and $EventLog ) {
		try {
			Write-msg -log -text "[EventLog]:got:[$($RemoteComputers.count)]"
			
			#region Iegūstam EventLog datus

			$RemoteComputers = $RemoteComputers | ForEach-Object { $_.tolower() }

			if ( $PSCmdlet.ParameterSetName -like "InPathEventLog" ) {

				$paramGetEvent = @{
					InPath = $InPath
				}
			}
			if ( $PSCmdlet.ParameterSetName -like "NameEventLog" ) {
				
				$paramGetEvent = @{
					ComputerName = $RemoteComputers
				}
			}

			# iestatam noklusētās 30 dienas, ja nav norādīts savādāk
			if ( $Days ) { $paramGetEvent.Add('Days', $Days) } else
			{ $paramGetEvent.Add('Days', 30) }
			$result = Get-CompEvent @paramGetEvent
			
			#endregion

			#region Parādam rezultātu ekrānā

			$paramShowEvent = @{
				InputObject = $result
				OutPath     = $OutPath
			}

			if ( $PSCmdlet.ParameterSetName -like "NameEventLog" ) {
				$paramShowEvent.Add('Named', $True)
			}

			if ($OutPath) {
				$paramShowEvent.Add('InPathFileName', "$((Resolve-Path $InPath).Path)")
			}
			Show-CompEvent @paramShowEvent

			#endregion

		}
		catch {
			Write-ErrorMsg -Name 'EventLog' -InputObject $_
		}
	}
	else {
		if ( $PSCmdlet.ParameterSetName -like "*EventLog" ) {
			Write-msg -log -bug -text "[EventLog] No computer in list."
		}
	}
	#endregion

	<# ---------------------------------------------------------------------------------------------------------
		Izpildām Install/Uninstall skriptu - $CompProgramFileName = "lib\Set-CompProgram.ps1"
		Izpildām WakeOnLan skriptu - $CompWakeOnLanFileName	= "lib\Invoke-CompWakeOnLan.ps1"
	--------------------------------------------------------------------------------------------------------- #>
	#region INSTALL, UNINSTALL
	if ( ( $RemoteComputers.count -gt 0 ) -and ( $Install -or $Uninstall -or $WakeOnLan ) ) {

		try {
			Write-msg -log -text "[$(if($Install){"Install"}elseif($Uninstall){"Uninstall"}else{"WakeOnLan"})]:got:[$($RemoteComputers.count)]"
			Write-Verbose "[$(if($Install){"Install"}elseif($Uninstall){"Uninstall"}else{"WakeOnLan"})]:got:[$($RemoteComputers.count)]"
			
			$parameterVSkJob = @{
				Computers = $RemoteComputers
			}

			if ( $PSCmdlet.ParameterSetName -eq "Name4Install" -or
				$PSCmdlet.ParameterSetName -eq "InPath4Install" ) {

				# Install-CompProgram.ps1 [-ComputerName] <string> [-InstallPath <FileInfo>] [<CommonParameters>]
				# Write-Host "[Installer] waiting for results:"

				#kriptējam $Install parametru
				Write-Verbose "[Main] Install.FullName [$($Install.FullName)]"
				# $secParameter = $Install.FullName | ConvertTo-SecureString -AsPlainText -Force
				# $EncryptedInstallPath = $secParameter | ConvertFrom-SecureString
				# Write-Verbose "[Main] EncryptedInstallPath [$EncryptedInstallPath]"

				$parameterVSkJob.Add('ScriptBlockName', 'Install')
				$parameterVSkJob.Add('InstallPath', $Install.FullName)
			}
			elseif ( $PSCmdlet.ParameterSetName -eq "Name4Uninstall" -or
				$PSCmdlet.ParameterSetName -eq "InPath4Uninstall" ) {

				#kriptējam $Uninstall parametru
				$secParameter = $Uninstall | ConvertTo-SecureString -AsPlainText -Force
				$EncryptedParameter = $secParameter | ConvertFrom-SecureString

				# Uninstall-CompProgram.ps1 [-ComputerName] <string> [-CryptedIdNumber <string>] [<CommonParameters>]
				Write-Host "[Uninstaller] waiting for results:"
				
				$parameterVSkJob.Add('ScriptBlockName', 'Uninstall')
				$parameterVSkJob.Add('UninstallEncryptedParameter', $EncryptedParameter)

				# $JobResults = Send-VSkJob  $RemoteComputers -ScriptBlockName 'SBUninstall' -UninstallEncryptedParameter $EncryptedParameter
			}
			elseif ( $PSCmdlet.ParameterSetName -eq "WakeOnLanName" -or
					$PSCmdlet.ParameterSetName -eq "WakeOnLanInPath"  ) {

				# Invoke-CompWakeOnLan.ps1 [-ComputerName] <string[]> [-DataArchiveFile] <FileInfo> [<CommonParameters>]
				Write-Host "[Waker] waiting for results:"

				$parameterVSkJob.Add('ScriptBlockName', 'WakeOnLan')
				$parameterVSkJob.Add('DataArchiveFile', $DataArchiveFile)
			}
			
			$JobResults = Send-VSkJob @parameterVSkJob

			Write-Host "`n[Results]=====================================================================================================" -ForegroundColor Yellow
			$JobResults | Sort-Object -Property Computer, id `
			| Format-Table Computer, Message -AutoSize  `
			| Out-String -Stream | Where-Object { $_ -ne "" } `
			| ForEach-Object { if ( $_ -match '.*SUCCESS.*' ) { Write-Host "$_" -ForegroundColor Green } `
					elseif ($_ -match '.*ERROR.*') { Write-Host "$_" -ForegroundColor Red } `
					elseif ($_ -match '.*WARN.*') { Write-Host "$_" -ForegroundColor Yellow } `
					else { Write-Host "$_" } }
			Write-Host "--------------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow

		}
		catch {
			Write-ErrorMsg -Name $(if ($Install) { "Install" }elseif ($Uninstall) { "Uninstall" }else { "WakeOnLan" }) -InputObject $_
		}
	}

	#endregion
}

END {
	Stop-Watch -Timer $scriptWatch -Name Script
}