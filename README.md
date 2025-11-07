# DomainTools
# AD Domain Inventory Testing Repo

# Active Directory Infrastructure Documentation Script - Enhanced Version 5

1. Run this on a Domain Controller with appropriate permissions using PowerShell ISE.

2. Exchange Module for on-prem will need to be imported on the Domain Controller yyou are running this script from.

3. Go to a On-Prem Exchange Server, open PowerShell ISE as an admin.

4. Open C:\Program Files\WindowsPowerShell\Modules and copy the change module folder from there to the Domain Controller "note the path" because you want to paste it in the same path on the Domain Controller.

5. Go back to powershell ise on the domain controller, run Import-Module -Name 'Exchange' #<---orwhatever the name is "sorry I forgot the precise module name"...but you get where I am going with this. 

6. Once imported do the same cmd but instead of Import-Module do a Install-Module and it should be done. 

7. Close PowerShell ISE and reopen it as a admin and the exchange module should bee ready to use.

8. Enjoy!!!
