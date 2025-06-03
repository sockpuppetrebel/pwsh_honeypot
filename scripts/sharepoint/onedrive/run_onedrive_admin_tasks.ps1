# Import the functions
. .\manage_onedrive_admins.ps1

# STEP 1: Configure your tenant variables
$AdminURL = "https://YOURTENANT-admin.sharepoint.com/"  # Replace YOURTENANT with your actual tenant name
$AdminName = "admin@yourdomain.com"                      # Your admin account email
$UserToRemove = "user@yourdomain.com"                    # The user you want to remove from all sites

# Example with real values (replace these):
# $AdminURL = "https://contoso-admin.sharepoint.com/"
# $AdminName = "admin@contoso.com"
# $UserToRemove = "john.doe@contoso.com"

# STEP 2: Run the functions as needed

# Option A: Add yourself as admin to all OneDrive sites (useful for cleanup tasks)
Add-SPOAdminToAllOneDrive -AdminURL $AdminURL -AdminName $AdminName

# Option B: Remove a specific user from all OneDrive sites
Remove-UserFromAllOneDrive -AdminURL $AdminURL -UserToRemove $UserToRemove

# Option C: Remove admin access when done (best practice)
Remove-SPOAdminFromAllOneDrive -AdminURL $AdminURL -AdminName $AdminName