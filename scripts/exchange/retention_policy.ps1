#create new policy tag
New-RetentionPolicyTag -Name "Delete after 6 months" `
  -Type All `
  -RetentionEnabled $true `
  -AgeLimitForRetention 180 `
  -RetentionAction DeleteAndAllowRecovery
 
#Create a policy that uses the tag
New-RetentionPolicy -Name "6-Month Delete Policy" -RetentionPolicyTagLinks "Delete after 6 months"

#assign the policy to mailbox
Set-Mailbox -Identity "user@optimizely.com" -RetentionPolicy "6-Month Delete Policy"

#(optional) verify assignment
Get-Mailbox -Identity "user@optimizely.com" | Select RetentionPolicy

#when does it take effect you ask?
# The Managed Folder Assistant processes mailboxes on a schedule.

Start-ManagedFolderAssistant -Identity "user@optimizely.com"

