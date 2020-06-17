	Import-csv C:\csv\UsersNoProxy.csv | ForEach-Object {
	$ID = $_.PrimarySmtpAddress
	$alias = $_.Alias
	$365Email = $alias + "@contoso.mail.onmicrosoft.com"
	Set-Mailbox -Identity $ID -EmailAddresses @{add = $365Email}
	}
