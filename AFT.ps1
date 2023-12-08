#Full Path to the AFT Log
$AFT_log = "C:\tools\Test_v2.xlsx"

#Function to see if a file is open.
Function Test-FileLock {
    param (
      [parameter(Mandatory=$true)][string]$Path
    )
  
    $oFile = New-Object System.IO.FileInfo $Path
  
    if ((Test-Path -Path $Path) -eq $false) {
      return write-host "File doesn't exist"
    }
  
    try {
      $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  
      if ($oStream) {
        $oStream.Close()
      }
      # File is not locked by a process
      # File is not open
      return $true
        
    } catch {
      # file is locked by a process.
      # File is open/in use
      return $false
       
    }
  }
  
#Gets AFT Number
#this just counts the number of folders and then adds 1, there isn't any validation yet.
$AFT_num = Get-ChildItem -Path C:\tools\Logs -Directory | Where-Object name -Like "AFT-2023-LS-*" | Select-Object name
$AFT_num = $AFT_num.name.count + 1 | ForEach-Object tostring AFT-2023-LS-000

#Asking user questions in a while loop, to verify information is correct.
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
$Description = Read-Host "Please enter the description"
$classification = Read-Host "Enter Classification of the data"
$classification = $classification.ToUpper()


$ans = "n"
While ($ans -ne "y") {
    #Verifying the User answers are correct
    write-host -ForegroundColor Green "`n`nAFT Control Number: $($AFT_num)"
    Write-host -ForegroundColor Green "DTA: $($DTA)"
    write-host -ForegroundColor Yellow "1) Folder Path: $($folder_path)"
    write-host -ForegroundColor Yellow "2) Additional Control Numbers: $($Other_controlNum)"
    write-host -ForegroundColor Yellow "3) Media Type: $($MediaType)"
    write-host -ForegroundColor Yellow "4) Source: $($source)"
    write-host -ForegroundColor Yellow "5) SME/Requstor: $($SME)"
    write-host -ForegroundColor Yellow "6) Description: $($Description)"
    write-host -ForegroundColor Yellow "7) Data Classification: $($classification)"
    
    $ans = Read-host -Prompt "Is this information correct? [Y/N]"
    While (($ans -ne "n") -and ($ans -ne "y")){
        $ans = Read-host -Prompt "Is this information correct? [Y/N]"
    }

    if ($ans -eq "n"){
        switch (Read-host -Prompt "which line do you want to edit? [1-7]")
        {
            1{$Folder_path = New-Object System.Windows.Forms.FolderBrowserDialog
                $Null = $Folder_path.ShowDialog()
                $folder_path = $folder_path.SelectedPath}
            2{$Other_controlNum = Read-Host -Prompt "Please enter additional control numbers (If applicable)"}
            3{$MediaType = Read-host -Prompt "Media Type (ex. DVD, Padlock...)"}
            4{$source = Read-host -Prompt "Source"}
            5{$SME = Read-Host -Prompt "Enter SME/Requestor's Name"}
            6{$Description = Read-Host "Please enter the description"}
            7{$classification = Read-Host "Enter Classification of the data"
                $classification= $classification.ToUpper()}
            Default{Write-host -ForegroundColor Red "`nLine not found. Please select 1-7"}
        }
    }

}

#Telling the user what is going on.
Write-host "`nGathering information on for selected folder."

#pulling the dates from the system
$date = get-date -Format yyyy-MM-dd

#Creates New Log-000 folder
#Need to test to see if the folder already exsits. 
New-Item -Path C:\tools\Logs\ -ItemType Directory -Name $AFT_num | out-null

#Gets list of files
Get-ChildItem -Recurse -path $folder_path | Out-File C:\tools\Logs\$AFT_num\$($AFT_num)_filelist.txt
$file_list = Get-ChildItem -path C:\tools\Logs\$AFT_num\$($AFT_num)_filelist.txt

#Gather File Type information
$extensions = Get-ChildItem -Path $folder_path -Recurse -file | ForEach-Object {$_.Extension.ToLower()} | sort-object | Get-Unique
$extensions = $extensions -join ' '
$extensions | Out-File "C:\tools\Logs\$($AFT_num)\$($AFT_num)_filelist.txt" -append

#Get Folder Size
$size = "{0:N2} GB" -f ((Get-ChildItem -force -Path $folder_path -Recurse | Measure-Object Length -Sum).sum /1GB)
#If the folder size is less than 1GB it will re-run it as MB
If ($size -lt "1"){
    $size = "{0:N2} MB" -f ((Get-ChildItem -force -Path $folder_path -Recurse | Measure-Object Length -Sum).sum /1MB)
}
$size | Out-File "C:\tools\Logs\$($AFT_num)\$($AFT_num)_filelist.txt" -append

Write-host "Updating the AFT Log.`n"

#Test to see if the AFT log is open/ in use. If its open/in use it writes the information to a csv file located in the AFT-log folder.
If (Test-FileLock $AFT_log){
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $wb = $excel.workbooks.open($AFT_log)
    $ws = $wb.worksheets.Item("Incoming")

    #$ws.UsedRange.Rows.Count
    #Find empty cell
    $row = 1
    Do {$row | out-null } until ($null -eq $ws.cells.Item($row++,1).value2)
    #Need to -1 to find the empty row
    $row = $row -1

    $ws.Cells.Item($row,1).Value="$($AFT_num)"
    $ws.Cells.Item($row,2).Value="$($Other_controlNum)"
    $ws.Cells.Item($row,3).Value="$($date)"
    $ws.Cells.Item($row,4).Value="$($MediaType)"
    $ws.Cells.Item($row,5).Value="$($source)"
    $ws.Cells.Item($row,6).Value="$($SME)"
    $ws.Cells.Item($row,7).Value="$($DTA)"
    $ws.Cells.Item($row,8).Value="$($Description)"
    $ws.Cells.Item($row,9).Value="Filenames: $($file_list.FullName)"
    $ws.Cells.Item($row,10).Value="$($extensions)"
    $ws.Cells.Item($row,11).Value="$($size)"
    $ws.Cells.Item($row,12).Value="$($classification)"

    $wb.save()
    $wb.close()
    $excel.quit()

    Write-Host "`n`nThe AFT Log has been updated. Please verify that it is correct"
    write-host "$AFT_log`n"
    Read-Host -Prompt "Press Enter to close this window" | Out-Null
}else{
    Write-Host -ForegroundColor Red "The AFT log is currently open!`n"
    write-host -ForegroundColor Yellow "Writing information to a CVS file."
    $csv = "c:\tools\CSV_$($AFT_num).csv"
    $header = "AFT Control Number","Other Inventory Number","Date","Media Type","Source","SME","DTA","Description","File names","File formats","Size","Classification"
    $header -join "," | Out-File -FilePath $csv -Encoding utf8
    $data = @($AFT_num,$Other_controlNum,$date,$MediaType,$source,$SME,$DTA,$Description,$file_list.FullName,$extensions,$size,$classification)
    $data -join "," | Add-Content -Path $csv
    write-host "Information has been written to $($csv)"
    Read-Host -Prompt "Press Enter to close this window" | Out-Null
}
<#
Columns in the Excel doc
1 AFT Control Number
2 Other Inventory Number (If applicable) 
3 Date
4 Media Type
5 Source
6 SME
7 DTA
8 Description
9 File names
10 File formats
11 Size
12 Classification
#>
