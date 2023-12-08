#Gets AFT Number
$AFT_num = Get-ChildItem -Path C:\tools\Logs -Directory | where name -Like "AFT-Log-FY2023-*" | select name
$AFT_num = $AFT_num.name.count + 1 | % tostring AFT-Log-FY2023-000


#Asking user questions in a while loop, to verify information is correct.
$ans = "n"
While ($ans -ne "y") {
    # Ask User questions
    write-host -ForegroundColor green "`n`nAFT Control Number: $($AFT_num)"
    Add-Type -AssemblyName System.Windows.Forms
    Write-host "Use the folder browser to select the folder:"
    
	#Folder picker popup
    $Folder_path = New-Object System.Windows.Forms.FolderBrowserDialog
    $Null = $Folder_path.ShowDialog()
    $folder_path = $folder_path.SelectedPath
    
    $Other_controlNum = Read-Host -Prompt "Please enter additional control numbers (If applicable)"
    $MediaType = Read-host -Prompt "Media Type (ex. DVD, Padlock...)"
    $source = Read-host -Prompt "Source"
    $SME = Read-Host -Prompt "Enter SME/Requestor's Name"
    $DTA = $env:UserName
    $DTA = $DTA.toupper()
	$Description = Read-Host "Please enter the description"
    $classification = Read-Host "Enter Classification"
    $classification = $classification.ToUpper()

    #Verifying the User answers are correct
    write-host -ForegroundColor Green "`n`nAFT Control Number: $($AFT_num)"
    write-host -ForegroundColor Yellow "Folder Path: $($folder_path)"
    write-host -ForegroundColor Yellow "Additional Control Numbers: $($Other_controlNum)"
    write-host -ForegroundColor Yellow "Media Type: $($MediaType)"
    write-host -ForegroundColor Yellow "Source: $($source)"
    write-host -ForegroundColor Yellow "SME/Requstor: $($SME)"
    Write-host -ForegroundColor Yellow "DTA: $($DTA)"
    write-host -ForegroundColor Yellow "Description: $($Description)"
    write-host -ForegroundColor Yellow "Classification: $($classification)"
    $ans = Read-host -Prompt "Is this information correct? [Y/N]"
    While (($ans -ne "n") -and ($ans -ne "y")){
        $ans = Read-host -Prompt "Is this information correct? [Y/N]"
    }
}

#Get current date
$date = get-date -Format yyyy-MM-dd

#Creates New Log-000 folder
#Need to test to see if the folder already exsits. 
New-Item -Path C:\tools\Logs\ -ItemType Directory -Name $AFT_num | out-null

#Gets list of files
$file_list = Get-ChildItem -Recurse -path $folder_path | Out-File C:\tools\Logs\$AFT_num\$($AFT_num)_filelist.txt

#Gather File Type information
$extensions = Get-ChildItem -Path $folder_path -Recurse -file | % {$_.Extension.ToLower()} | unique
$extensions = $extensions -join ' '
$extensions | Out-File "C:\tools\Logs\$($AFT_num)\$($AFT_num)_filelist.txt" -append



#Get Folder Size
$size = "{0:N2} GB" -f ((Get-ChildItem -force -Path $folder_path -Recurse | measure Length -Sum).sum /1GB)
If ($size -lt "1"){
    $size = "{0:N2} MB" -f ((Get-ChildItem -force -Path $folder_path -Recurse | measure Length -Sum).sum /1MB)
}
$size | Out-File "C:\tools\Logs\$($AFT_num)\$($AFT_num)_filelist.txt" -append

$AFT_log = "C:\tools\Test_v2.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.workbooks.open($AFT_log)
$ws = $wb.worksheets.Item("Incoming")

#Find empty cell
$row = 1
Do {$row | out-null } until ($ws.cells.Item($row++,1).value2 -eq $null)
#Need to -1 to find the first empty row
$row = $row -1

$ws.Cells.Item($row,1).Value="$($AFT_num)"
$ws.Cells.Item($row,2).Value="$($Other_controlNum)"
$ws.Cells.Item($row,3).Value="$($date)"
$ws.Cells.Item($row,4).Value="$($MediaType)"
$ws.Cells.Item($row,5).Value="$($source)"
$ws.Cells.Item($row,6).Value="$($SME)"
$ws.Cells.Item($row,7).Value="$($DTA)"
$ws.Cells.Item($row,8).Value="$($Description)"
$ws.Cells.Item($row,9).Value="Filenames: C:\tools\Logs\$($AFT_num)\$($AFT_num)_filelist.txt"
$ws.Cells.Item($row,10).Value="$($extensions)"
$ws.Cells.Item($row,11).Value="$($size)"
$ws.Cells.Item($row,12).Value="$($classification)"

$wb.save()
$wb.close()
$excel.quit()
