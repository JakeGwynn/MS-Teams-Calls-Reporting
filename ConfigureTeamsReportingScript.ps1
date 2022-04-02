$AuthFileFolder = "C:\Temp\TeamsPBI"

$SecureClientSecret = Read-Host -Prompt "Enter Client Secret" -assecurestring
$EncryptedClientSecret = ConvertFrom-SecureString $SecureClientSecret 
$EncryptedClientSecret > "$AuthFileFolder\ClientSecret.txt"

$SecureAdminPassword = Read-Host -Prompt "Enter Service Account password" -assecurestring
$EncryptedAdminPassword = ConvertFrom-SecureString $SecureAdminPassword 
$EncryptedAdminPassword > "$AuthFileFolder\AdminPassword.txt"

$SecureSmtpPassword = Read-Host -Prompt "Enter SMTP account password" -assecurestring
$EncryptedSmtpPassword = ConvertFrom-SecureString $SecureSmtpPassword 
$EncryptedSmtpPassword > "$AuthFileFolder\SmtpPassword.txt"
