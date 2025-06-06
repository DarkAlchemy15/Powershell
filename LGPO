$gpos = Get-ChildItem -Directory -Path "C:\tools\Test"

$log = @()
foreach($line in $gpos){
    $row = New-Object psobject

    $gpo_name = (select-xml -Path ".\Test\$($line.Name)\bkupinfo.xml" -XPath //*).Node.GPODisplayName."#cdata-section"
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

