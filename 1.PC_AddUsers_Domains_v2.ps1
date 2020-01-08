## CONNECTING TO PARNER CENTER:
Import-Module MSOnline

#Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir script is running 
$PCDomains_Csv = Read-Host -Prompt "Enter the location for the csv file containing PC Domains List: (PCDomains_EMEA.csv)" 
$PCDomains_Csv = $PSScriptRoot + '\' + $PCDomains_Csv
$UserstoAdd_Csv = Read-Host -Prompt "Enter the location for the csv file containing Users List: (UserstoAddPC.csv)"
$UserstoAdd_Csv = $PSScriptRoot + '\' + $UserstoAdd_Csv

#######################################################################################################################
#######################################################################################################################
# Onboard users to all PC # Partner Center login iterations for connection ############################################
#######################################################################################################################

Import-Csv -Path $PCDomains_Csv | 
 ForEach-Object { `
    $FullUser = $SecUser + $_.MailDomain.trim()
    $MailDomain = $_.MailDomain.trim()
    $UsageLocation = $_.UsageLocation.trim()
    
    Write-Host "`n- Login Credentials Used for Partner Center:" $FullUser "`n"
        
    # Conexion to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        Import-Csv -Path $UserstoAdd_Csv |
            foreach {
                #Clear pasted spaces
                $FirstName = $_.FirstName.trim()
                $LastName = $_.LastName.trim()
                $AdminGroup = $_.AdminGroup.trim()
                #$Position = $_.Position.trim() #No trimming spaces
                
                $UPN = $FirstName + "." + $LastName + $MailDomain
                $DisplayName = $FirstName + " " + $LastName + " " + $_.Position
                #Write-Host "`n- VAriables" $FirstName + $LastName + $AdminGroup + $UPN + $DisplayName "`n"

                $UserExists = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue

                If ($UserExists)  {
                Write-Host "User $UPN already exists at:" $FullUser "`n"
                } Else {
                New-MsolUser `
                    -UserPrincipalName $UPN `
                    -FirstName $FirstName `
                    -LastName $LastName `
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
                    -PhoneNumber '0844 494 4480' 
                
                # REview getting accounts to be sent out
                #| Export-Csv -Path $PSScriptRoot + '\UsersAccounts.csv' -NoTypeInformation -Append
                
                # Add Agent role. Obtain UserID for Admin agents group ("AdminAgents","HelpdeskAgents","SalesAgents")
                $GroupID = Get-MsolGroup | Where-Object {$_.DisplayName -eq $AdminGroup}
                # Check users correct memberships
                # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                
                $User = Get-MsolUser -UserPrincipalName $UPN
                Add-MsolGroupMember `
                    -GroupObjectId $GroupID.ObjectId `
                    -GroupMemberType 'User' `
                    -GroupMemberObjectId $User.ObjectId #-ErrorAction Inquire
                
                # Add Admin role #Get-MsolRole -RoleName 'User Account Administrator' | Sort Name | Select Name,Description
                # Billing Administrator         Can perform common billing related tasks like updating payment information.
                # User Account Administrator    Can manage all aspects of users and groups, including resetting passwords for limited admins.
                    If ($_.RoleName.trim()) {
                       Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $_.RoleName -ErrorAction SilentlyContinue
                    }
                }
            }
}