#status
Get-MailboxDatabase -Server elefante2 -Status | Format-List mounted

#size
Get-MailboxDatabase -Status| select Databasesize
