#Create Shared Mailbox
$NameSharedMailbox = Read-Host -Prompt “Successfactors.support@optimizely.com”
New-Mailbox -Shared -Name $NameSharedMailbox -DisplayName $NameSharedMailbox -Alias $NameSharedMailbox
$ShareMailboxGetGUID = Get-Mailbox -Identity $NameSharedMailbox
#Pick Deleted Mailbox To Restore
$mailboxtoRestore = Get-Mailbox -SoftDeletedMailbox | Select Name,PrimarySMTPAddress,DistinguishedName | Sort-Object Name | Out-GridView -Title “Current Softdeleted Mailbox List” -PassThru
#Command to restore deleted mailbox to new shared mailbox.
New-MailboxRestoreRequest -SourceMailbox $mailboxtoRestore.DistinguishedName -TargetMailbox $ShareMailboxGetGUID.PrimarySmtpAddress -AllowLegacyDNMismatch
