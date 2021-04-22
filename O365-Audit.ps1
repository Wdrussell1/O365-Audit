#Created By Casey Barrett
#For the purpose of gathering data about Administrator Accounts, Domains and their status, and the MFA Status for each user.
#Nothing in this script is considered the property of anyone
#The commands are simple input/outputs and designed for ease of use. 
#
#
#
#
#
#
#Import Export-Excel
write-output "Installing ImportExcel Module. Visit https://github.com/dfinke/ImportExcel and https://www.powershellgallery.com/packages/ImportExcel/7.1.1 for more information." 
Install-Module ImportExcel -Repository PsGallery -Force -AllowClobber

write-output "Be sure you are connected to the O365 domain before running this script! Any errors beyond this point are likely related to that if you get any."

#Main Script Start

$CompanyName = Read-Host -Prompt 'Input the company name'


if(!(Test-Path $CompanyName -PathType Container)) { 
    write-host "$CompanyName Path not found"
	write-host "Creating $CompanyName"
	mkdir C:/$CompanyName
} else {

}

#Getting Administrator lists

Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Company Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Company Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "User Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "User Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Service Support Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Service Support Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Directory Readers").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Directory Readers" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Exchange Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Exchange Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "SharePoint Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "SharePoint Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Azure AD Joined Device Local Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Azure AD Joined Device Local Ad" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Directory Synchronization Accounts").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Directory Synchronization Accou" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Application Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Application Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Application Developer").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Application Developer" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Cloud Device Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Cloud Device Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Authentication Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Authentication Administrator" -path C:\$CompanyName\Administrators.xlxs
Get-MsolRoleMember -RoleObjectId $(Get-MsolRole -RoleName "Teams Administrator").ObjectId | select DisplayName,EmailAddress,IsLicensed,RoleMemberType | Export-Excel -workSheetName "Teams Administrator" -path C:\$CompanyName\Administrators.xlxs

#Cleaning up the Excel Workbook

Join-Worksheet C:\$CompanyName\Administrators.xlxs  -workSheetName "Admin Users" -AutoFilter -AutoSize -FromLabel "Account Type"

Remove-Worksheet -workSheetName "Company Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "User Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Service Support Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Directory Readers" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Exchange Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "SharePoint Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Azure AD Joined Device Local Ad" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Directory Synchronization Accou" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Application Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Application Developer" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Cloud Device Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Authentication Administrator" -path C:\$CompanyName\Administrators.xlxs
Remove-Worksheet -workSheetName "Teams Administrator" -path C:\$CompanyName\Administrators.xlxs

#Listing Domains

Get-MsolDomain | select Name,Authentication,Capabilities,Status | Export-Excel -workSheetName "Domains" -path C:\$CompanyName\Domains.xlxs -AutoSize -AutoFilter

#Getting MFA Status

Get-MsolUser -all | select DisplayName,UserPrincipalName,@{N="MFA Status"; E={ if( $_.StrongAuthenticationMethods.IsDefault -eq $true) {($_.StrongAuthenticationMethods | Where IsDefault -eq $True).MethodType} else { "Disabled"}}} | export-excel -workSheetName "MFAStatus" -path C:\$CompanyName\MFAStatus.xlxs -AutoSize -AutoFilter

#Data Consolidation

Copy-ExcelWorksheet -SourceObject C:\$CompanyName\Domains.xlxs -Sourceworksheet "Domains" -DestinationWorkbook C:\$CompanyName\O365-Report.xlxs -DestinationWorksheet "Domains"
Copy-ExcelWorksheet -SourceObject C:\$CompanyName\Administrators.xlxs -Sourceworksheet "Admin Users" -DestinationWorkbook C:\$CompanyName\O365-Report.xlxs -DestinationWorksheet "Admin Users"
Copy-ExcelWorksheet -SourceObject C:\$CompanyName\MFAStatus.xlxs -Sourceworksheet "MFAStatus" -DestinationWorkbook C:\$CompanyName\O365-Report.xlxs -DestinationWorksheet "MFAStatus"


write-output "Data collection complete. Please see C:/$CompanyName/ for all related files. O365-Report.xlxs is the consolidated data, the other files are reference files."



