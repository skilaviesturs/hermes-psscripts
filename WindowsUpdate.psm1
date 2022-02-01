#Requires -Version 5.1
#Dot source all functions in all ps1 files located in the module folder
Get-ChildItem -Path "$PSScriptRoot\modules\privet\*.ps1" -Exclude Temporary.ps1 |
ForEach-Object {
    . $_.FullName
}
Get-ChildItem -Path "$PSScriptRoot\modules\*.ps1" -Exclude Temporary.ps1 |
ForEach-Object {
    . $_.FullName
}