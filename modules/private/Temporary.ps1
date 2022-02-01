
#region Get-NormaliseDiskLabelsForExcel
Function Get-NormaliseDiskLabelsForExcel {
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Computers
    )
    $TemplateDisks = [PSCustomObject][ordered]@{}
    #Get uniqe disk label's collection from computers and write in array
    foreach ($computer in $computers) {
        foreach ( $Item in $computer.PsObject.Properties ) {
            if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize' ) {
                if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
                    $TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
                }
            }
            if ( $Item.Name -match '^hDisk\s[a-zA-Z]:\ssize\sFree' ) {
                if ( -not ( Get-Member -InputObject $TemplateDisks -Name $Item.Name ) ) {
                    $TemplateDisks | Add-Member -MemberType NoteProperty -Name $Item.Name -Value 'none'
                }
            }
        }
    }
    #Add to each computer's properties missing disk label
    foreach ( $computer in $computers ) {
        foreach ($label in $TemplateDisks.PsObject.Properties) {
            if ( -not ( Get-Member -InputObject $computer -Name $label.Name ) ) {
                Add-Member -InputObject $computer -NotePropertyName $label.Name -NotePropertyValue 'none' -Force
            }
        }
    }

    $computers
}
#endregion

