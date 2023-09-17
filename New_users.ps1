<#
    .DESCRIPTION
        Create New users 1 at a time for multiple with a CSV file
        The CSV file only needs 3 Colums "FirstName, LastName, Description"
    .EXAMPLE
        C:\scripts\New_users.ps1 -file c:\scripts\Users.csv
    
#>
param(
    [string]$file
)

## List of Global Varibles
# List all default Usergroups
$Group_List = @(
    'Group_1'
    'Group_2'
)

# Single User Creation, if no file is
if (-not $PSBoundParameters.ContainsKey('file')){
    $FirstName = Read-host -Prompt "Enter First Name: "
    $LastName = Read-Host -Prompt "Enter Last Name: "
    $Description = Read-Host -Prompt "Enter Description: "
    $Username = $LastName.ToLower() + $FirstName.ToLower()[0]
    $email = $Username + "@lab.local"
    $OU = "OU=Lab_Users,DC=Lab,DC=Local"
    $password = ConvertTo-SecureString "!QAZ" -AsPlainText -Force
    
    #Fixing First and Last Name so they Match
    # Ex. ToNy ==> Tony
    $FirstName = $Firstname.ToLower()
    $FirstName = $FirstName.Replace($FirstName[0],$FirstName[0].ToString().ToUpper())
    $LastName = $LastName.ToLower()
    $LastName = $LastName.Replace($LastName[0],$LastName[0].ToString().ToUpper())
    
    $New_User = @(
        Name = $($FirstName + " " + $LastName)
        GivenName = $FirstName
        Surname = $LastName
        EmailAddress = $email
        SamAccountName = $Username
        UserPrincipanName = $email
        AccountPassword = $password
        Enabled = $True
        Path = $OU
        Description = $Description
        ChangePasswordAtLogon = $True
    )

    New-ADUser @New_User
    $Group_List.foreach({Add-ADGroupMember -Identity $_ -Members $Username})
}

If ($PSBoundParameters.ContainsKey('file')){
    while(-not (Test-path $file -PathType Leaf )){
        Write-host -ForegroundColor Yellow "$($file) doesn't exist"
        $file = Read-Host -Prompt "Enter full file path: "
    }
    
    $user_list = import-csv -Path $file

    $user_list.foreach{
        New-ADUser `
        -Name $($_.Firstname + " " + $_.LastName) `
        -GivenName $_.FirstName `
        -Surname $_.LastName `
        -EmailAddress $($_.LastName.ToLower() + $_.FirstName.ToLower()[0] + "@lab.local") `
        -SamAccountName $($_.LastName.ToLower() + $_.FirstName.ToLower()[0]) `
        -UserPrincipalName $($_.LastName.ToLower() + $_.FirstName.ToLower()[0] + "@lab.local") `
        -Description $_.Description `
        -AccountPassword $(ConvertTo-SecureString "password123" -AsPlainText -Force) `
        -Path "OU=Lab_Users, DC=Lab,DC=local" `
        -Enabled $True
        -ChangePasswordAtLogon $True
    }

    $Usernames = $user_list.foreach({$_.LastName + $_.FirstName[0]})
    $Group_List.foreach({Add-ADGroupMember -Identity $_ -Members $Usernames})
}
