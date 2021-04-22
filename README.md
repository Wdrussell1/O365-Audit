# O365-Audit
This script is to audit O365 for a few details. This may evolve over time but it is designed as is.

The script will download and install Import Excel. You can find more information at https://github.com/dfinke/ImportExcel or https://www.powershellgallery.com/packages/ImportExcel/7.1.1.

This script will currently:
1. Gather the MFA status of all users and display only the ones not enabled fully.
2. Get all domains that are attached to the tenant.
3. Get all admin accounts and their roles (this does include built in accounts/roles)

You will need to connect to O365 before running this script. As it does not do this (as MFA is unable to be cleanly automated). I also did not want to store credentials in variables for obvious reasons.

Please read where the script is going to be putting information. It will be created in the root of C:/ inside its own folder. 

Might clean this readme up later but this is just a mechanism to share something simple i created. 
