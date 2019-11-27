## CONNECTING TO PARNER CENTER:
Import-Module MSOnline

## Write-Host "Details for Partner Center Credentials (same password and login needed for all accounts `
## if MFA enabled please allow some time for the script to complete)..."
## * # Un-Comment next three lines comments plus $O365Cred below comments for user and password capture and no prompt # * #
#$SecUser = Read-Host -Prompt "Enter your partner center account username (before the @) "
#$SecPass = Read-Host -Prompt "Enter your partner center account password "
#$SecPass = ConvertTo-SecureString $SecPass -AsPlainText -Force 

#Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir script is running 
$PCDomains_Csv = Read-Host -Prompt "Enter the location for the csv file containing PC Domains List: (PCDomains_EMEA.csv)" 
$PCDomains_Csv = $PSScriptRoot + '\' + $PCDomains_Csv
$UserstoAdd_Csv = Read-Host -Prompt "Enter the location for the csv file containing Users List: (UserstoAddPC.csv)"
$UserstoAdd_Csv = $PSScriptRoot + '\' + $UserstoAdd_Csv

#######################################################################################################################
#######################################################################################################################
# Onboard users to all PC # Partner Center login iterations for connection ############################################
#######################################################################################################################

Try {
Import-Csv -Path $PCDomains_Csv | 
 ForEach-Object { `
    Try { 
    $FullUser = $SecUser + $_.Maildomain
    $Maildomain = $_.Maildomain
    $UsageLocationPC = $_.UsageLocation
    
    ## * # Un-Comment next two lines for user and password no prompt for every PC Account # * #
    Write-Host "`n- Login Credentials Used for Partner Center:" $FullUser "`n"
    #$O365Cred = New-Object System.Management.Automation.PSCredential ($FullUser , $SecPass)
    
    # Conexion to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        Import-Csv -Path $UserstoAdd_Csv |
            foreach {
                Try { 
                $UPN = $_.Firstname + "." + $_.Lastname + $Maildomain
                $DisplayName = $_.Firstname + " " + $_.Lastname + " " + $_.Position #" [EMEA IMS L1]"
                $AdminGroup = $_.AdminGroup

                New-MsolUser `
                    -UserPrincipalName $UPN `
                    -FirstName $_.FirstName `
                    -LastName $_.LastName `
                    -DisplayName $DisplayName `
                    -Department 'Managed Services' `
                    -Title 'Analyst' `
                    -Password 'P@$$w0rd123' `
                    -UserType 'Member' `
                    -StreetAddress 'EMEA' `
                    -Office 'Managed Services EMEA' `
                    -City 'EMEA' `
                    -State 'EMEA' `
                    -Country $UsageLocation `
                    -PostalCode '28082' `
                    -PhoneNumber '0844 494 4480' -ErrorAction Inquire
                
                # Add Agent role (*ADD prompt for information)
                # Obtain UserID for Admin agents group ("AdminAgents","HelpdeskAgents","SalesAgents")
                #$GroupID = Get-MsolGroup | Where-Object { $_.DisplayName -eq “AdminAgents”}
				$GroupID = Get-MsolGroup | Where-Object { $_.DisplayName -eq $AdminGroup}
                # Check users correct memberships
                # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                
                $User = Get-MsolUser -UserPrincipalName $UPN
                Add-MsolGroupMember `
                    -GroupObjectId $GroupID.ObjectId `
                    -GroupMemberType 'User' `
                    -GroupMemberObjectId $User.ObjectId #-ErrorAction Inquire
                
                # Add Admin role (*ADD prompt for informtion)
                # Get-MsolRole -RoleName 'User Account Administrator' | Sort Name | Select Name,Description
                # Billing Administrator         Can perform common billing related tasks like updating payment information.
                # User Account Administrator    Can manage all aspects of users and groups, including resetting passwords for limited admins.
                Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $_.RoleName -ErrorAction Inquire
                    } Catch {
                    Write-Host "`n ## Error adding user or roles from CSV"
                    $_.Exception.Message
                    }
            }
    } Catch {
    Write-Host "`n ## Error Capturing Users CSV"
    $_.Exception.Message
    }
  }
} Catch {
Write-Host "`n ## Error Capturing Domain List CSV"
$_.Exception.Message
}