#clearing the output window
Clear-Host

#checking if 'ActiveDirectory' module is installed
$ErrorActionPreference = "Stop"
try {
    Import-Module -Name ActiveDirectory
} catch {
    Write-Host "'ActiveDirectory' module is not installed.  Terminating script..." -ForegroundColor Red
    Start-Sleep -Seconds 10
    exit
}

#setting error action to 'SilentlyContinue' so errors can be handled by the code instead of crashing the script or using a lot of try/catch blocks
$ErrorActionPreference = "SilentlyContinue"

#getting current folder of the script
$path = Split-Path -Path $MyInvocation.MyCommand.Path

#the timestamped output file will be saved in the same folder as the script
$outputPath = "$path\AdUsers.$(Get-Date -Format 'yyyyMMdd.HHmmss').csv"

#clearing previous variable values
$domain, $ouName, $groupName, $userName, $email, $emailServer, $emailTo, $emailFrom = $null, $null, $null, $null, $null, $null, $null, $null

#getting user input
$domain = Read-Host -Prompt "Enter target domain (or leave blank to query current domain)"
$ouName = Read-Host -Prompt "Enter OU name  (or leave blank to include users in any OU)"
$groupName = Read-Host -Prompt "Enter AD group name (or leave blank to include users in any group)"
$userName = Read-Host -Prompt "Enter AD user name (or leave blank to include all users)"
$email = Read-Host -Prompt "Would you like the CSV emailed (y/n)"
if ($email -eq "y") {
    $emailServer = Read-Host -Prompt "Enter SMTP server IP"
    $emailTo = Read-Host -Prompt "Enter list of 'to' email addresses, delimited by commas"
    $emailFrom = Read-Host -Prompt "Enter the 'from' email address (ex: automation@yourcompany.com)"
}

#getting current domain if a target domain was not entered
if ($domain.Length -eq 0) {
    $domain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
}

#uppercasing the domain so the case-sensitive Replace() method will work later
$domain = $domain.ToUpper()

#three different types of AD queries based on whether a user name or OU was entered
if ($userName.Length -gt 0) {
    $users = Get-ADUser -Server $domain -Identity $userName –Properties "msDS-UserPasswordExpiryTimeComputed", *
} else {
    $users = Get-ADUser -Server $domain -Filter {(Enabled -eq $true) -and (PasswordNeverExpires -eq $false)} –Properties "msDS-UserPasswordExpiryTimeComputed", *
}

#creating the output
$output = @()
foreach ($user in $users) {
    $output += "" | Select-Object -Property `
    @{n="DisplayName";e={$user.DisplayName}}, 
    @{n="SamAccountName";e={$user.SamAccountName}}, 
    @{n="Created";e={$user.Created}}, 
    @{n="PasswordLastSet";e={$user.PasswordLastSet}}, 
    @{n="PasswordExpiration";e={[datetime]::FromFileTime($user."msDS-UserPasswordExpiryTimeComputed")}}, 
    @{n="LastLogonDate";e={$user.LastLogonDate}}, 
    @{n="OU";e={(($user.CanonicalName).ToUpper()).Replace("$domain/", "").Replace("/$($user.CN)", "")}}, 
    @{n="MemberOf";e={@($user.MemberOf | Sort-Object | foreach {((($_ -split "CN=")[1]).Split(","))[0]}) -join ", "}}
}

#if an OU name was entered and a user name was not, this will filter the output to only include users within that OU
if (($ouName.Length -gt 0) -and ($userName.Length -eq 0)) {
    $output = $output | Where-Object {$_.OU -match $ouName}
}

#if a group name was entered and a user name was not, this will filter the output to only include users that are a member of that group
if (($groupName.Length -gt 0) -and ($userName.Length -eq 0)) {
    $output = $output | Where-Object {$_.MemberOf -match $groupName}
}

#validating output contains some data
if (($null -ne $output) -and (@($output).Count -gt 0)) {
    #sorting output so users with passwords expiring soon will be at the top
    $output = $output | Sort-Object PasswordExpiration
    $output | Export-Csv -Path $outputPath -NoTypeInformation -Force
    Write-Host "CSV file saved: $outputPath" -ForegroundColor Green
    if ($email -eq "y") {
        #attempting to send email using the email settings provided by user
        $ErrorActionPreference = "Stop"
        try {
            Send-MailMessage -Attachments $outputPath -SmtpServer $emailServer -To @(($emailTo -replace " ", "") -split ",") -From $emailFrom -Subject "AD Users ($domain)" -Body "Report attached."
            Write-Host "Report sent to $emailTo." -ForegroundColor Green
        } catch {
            Write-Host "Report NOT sent to $emailTo.  Please check email settings." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "No users found.  Please check input.  Terminating script..." -ForegroundColor Red
    Start-Sleep -Seconds 30
    exit
}

Write-Host "Script complete.  This window will close in 5 seconds..." -ForegroundColor Green
Start-Sleep -Seconds 5
exit
