<#
    .DESCRIPTION
        Use LGPO.exe to import multiple Local Policies 
    .NOTES
        Place GPOs to import in c:\tools\GPOs
    .SYNOPSIS
        Use LGPO.exe to import multiple Local polices. Place LGPO.exe and GPO's to import in 'C:\tools\GPOs'
#>

If (-Not (Test-Path c:\tools\GPOs)){
    Write-host -ForegroundColor Red "Place LGPO.exe and GPO's to import in 'C:\tools\GPOs'"
    exit
}

$path = "C:\tools\GPOs"
$gpos = Get-ChildItem -Directory -Path $path
$hostname = HOSTNAME
$date = Get-Date -Format yyyy-MM-dd_HHmm

If (-Not (Test-Path -Path $($path)\LGPO.exe)){
    Write-host -ForegroundColor Red "`nDid not find LGPO.exe"
    Write-host "Place LGPO.exe in the same folder as the GPOs`n"
    exit
}

$log = @()
$log += $hostname
$log += $date

foreach ($line in $gpos){
    $row = New-Object psobject

    $gpo_name = (select-xml -Path "$($path)\$($line.Name)\bkupinfo.xml" -XPath //*).Node.GPODisplayName."#cdata-section"
    $row | Add-Member -name "GPO_Name" -Value $($gpo_name)  -MemberType NoteProperty
    #(select-xml -Path ".\Test\$($line.Name)\bkupinfo.xml" -XPath //*).Node.GPODisplayName
    $row | Add-Member -name "GPO_ID" -Value $($line.Name) -MemberType NoteProperty

    $row | Out-Host
    Write-host -ForegroundColor Yellow "Importing $($gpo_name)"
    Write-host -foregroundcolor yellow "Running..... .\LGPO.exe \g $($line.Name)`n"
    
    $log += $row
    $row = $null
    
}

$log | Out-File -NoClobber -FilePath c:\tools\log.txt  -Append
Write-host -ForegroundColor Magenta "Finished.... Log file c:\tools\log.txt"

