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

#letting the user know this next command could take a few minutes to finish
Write-Host "`nQuerying the $domain domain for users matching input fields.  This may take a few minutes..." -ForegroundColor Green

$start = Get-Date

#two different types of AD queries based on whether a user name was entered
$users = @()
if ($userName.Length -gt 0) {
    $users += Get-ADUser -Server $domain -Identity $userName –Properties "msDS-UserPasswordExpiryTimeComputed", DisplayName, SamAccountName, Created, PasswordLastSet, LastLogonDate, CanonicalName, CN, MemberOf | Select-Object "msDS-UserPasswordExpiryTimeComputed", DisplayName, SamAccountName, Created, PasswordLastSet, LastLogonDate, CanonicalName, CN, MemberOf
} else {
    $users += Get-ADUser -Server $domain -Filter {(Enabled -eq $true) -and (PasswordNeverExpires -eq $false)} –Properties "msDS-UserPasswordExpiryTimeComputed", DisplayName, SamAccountName, Created, PasswordLastSet, LastLogonDate, CanonicalName, CN, MemberOf | Select-Object "msDS-UserPasswordExpiryTimeComputed", DisplayName, SamAccountName, Created, PasswordLastSet, LastLogonDate, CanonicalName, CN, MemberOf
}

$end = Get-Date
$duration = [int]((New-TimeSpan -Start $start -End $end).TotalSeconds)

Write-Host "`n$(@($users).Count) users found in AD ($duration seconds)." -ForegroundColor Cyan

Write-Host "`nAppending to output variable..." -ForegroundColor Green
$start = Get-Date

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

$end = Get-Date
$duration = [int]((New-TimeSpan -Start $start -End $end).TotalSeconds)
Write-Host "`n$(@($output).Count) users appended to output variable ($duration seconds)." -ForegroundColor Cyan

#if an OU name was entered and a user name was not, this will filter the output to only include users within that OU
if (($ouName.Length -gt 0) -and ($userName.Length -eq 0)) {
    $start = Get-Date
    $output = $output | Where-Object {$_.OU -match $ouName}
    $end = Get-Date
    $duration = [int]((New-TimeSpan -Start $start -End $end).TotalSeconds)
    Write-Host "`n$(@($output).Count) users included after applying OU filter ($duration seconds)." -ForegroundColor Cyan
}

#if a group name was entered and a user name was not, this will filter the output to only include users that are a member of that group
if (($groupName.Length -gt 0) -and ($userName.Length -eq 0)) {
    $start = Get-Date
    $output = $output | Where-Object {$_.MemberOf -match $groupName}
    $end = Get-Date
    $duration = [int]((New-TimeSpan -Start $start -End $end).TotalSeconds)
    Write-Host "`n$(@($output).Count) users included after applying group filter ($duration seconds)." -ForegroundColor Cyan
}

#validating output contains some data
if (($null -ne $output) -and (@($output).Count -gt 0)) {
    #sorting output so users with passwords expiring soon will be at the top
    $output = $output | Sort-Object PasswordExpiration
    $output | Export-Csv -Path $outputPath -NoTypeInformation -Force
    Write-Host "`nCSV file saved: $outputPath" -ForegroundColor Green
    if ($email -eq "y") {
        #attempting to send email using the email settings provided by user
        $ErrorActionPreference = "Stop"
        try {
            Send-MailMessage -Attachments $outputPath -SmtpServer $emailServer -To @(($emailTo -replace " ", "") -split ",") -From $emailFrom -Subject "AD Users ($domain)" -Body "Report attached."
            Write-Host "`nReport sent to $emailTo." -ForegroundColor Green
        } catch {
            Write-Host "`nReport NOT sent to $emailTo.  Please check email settings." -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "`nNo users found.  Please check input.  Terminating script..." -ForegroundColor Red
    Start-Sleep -Seconds 30
    exit
}

Write-Host "`nScript complete.  This window will close in 5 seconds..." -ForegroundColor Green
Start-Sleep -Seconds 5
exit
