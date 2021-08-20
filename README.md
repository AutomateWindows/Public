Get-ADUser -SearchBase "OU=Teachers,OU=Administration,OU=Users,OU=School-ETN,DC=na,DC=MYDOMAIN,DC=com" -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")} >> C:\temp\test2_SL.csv
