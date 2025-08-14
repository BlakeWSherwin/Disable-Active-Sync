# Disable-Active-Sync
Disable Active Sync with logging

This script will check all domain accounts for membership in the "ActiveSync-Access-Allowed" security group. It will then disable Exchange ActiveSync for any user who is not a member. This can also be called as a scheduled task to run every 15 or 30 minutes.

This will also send an email summary of all changes made. If you dont wish to use the email function simply change $true to $false on line 9 after $SendEmail
