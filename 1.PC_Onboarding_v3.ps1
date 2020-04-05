#######################################################################################################################
#######################################################################################################################
## PARNER CENTER AUTOMATION INSIGHT ###################################################################################
#######################################################################################################################
#######################################################################################################################

Import-Module MSOnline

#######################################################################################################################
#######################################################################################################################
# Onboard users to all PC # Partner Center login iterations for connection ############################################
#######################################################################################################################
# Global Variables Definitions ########################################################################################
# Hardcoded Login Credentials (Needs password match among all Partner Center acccounts ################################
#######################################################################################################################
# $global:SecUser = Read-Host -Prompt "Enter your partner center account username (before the @) "
# $global:SecPass = Read-Host -Prompt "Enter your partner center account password "
# $global:SecPass = ConvertTo-SecureString $SecPass -AsPlainText -Force

########################################################################################################################
# Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir where script is running #####
########################################################################################################################
# $global:PCDomains_Csv = $null
$global:PCUsers_Csv = $null

########################################################################################################################
# Path for storing output documents ####################################################################################
########################################################################################################################
$global:Date = $null
$global:CSVSavePath = $null


Function UpdateFilesCaptured {
param (
           [string]$Title = 'Partner Center Onboarding'
     )
     cls
     Write-Host "================ $Title ================"
#Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir script is running 
#$global:PCDomains_Csv = Read-Host -Prompt "`nEnter the location for the csv file containing PC Domains List (PCDomains.csv)" 
#$global:PCDomains_Csv = $PSScriptRoot + '\' + $PCDomains_Csv
$global:PCUsers_Csv = Read-Host -Prompt "`nEnter the location for the xlsx file containing Users List (PCUsers.xlsx)"
$global:PCUsers_Csv = $PSScriptRoot + '\' + $PCUsers_Csv
#Path for storing output documents
$global:Date = Get-Date -Format "d_MM_yyyy"
$global:CSVSavePath = $PSScriptRoot + '\' + $Date
Write-Host "`n"
}

Function PrepareFolders{

    Write-Host -ForegroundColor Yellow "`nScript now validating and preparing folder structure for output ..."
    
    If ((Test-Path $CSVSavePath) -eq $false){
        New-Item -ItemType Directory $CSVSavePath | Out-Null
        Write-Host -ForegroundColor Green "Folder" $CSVSavePath "has been created`n"
        }
    elseif((Test-Path $CSVSavePath) -eq $true){
        Write-Host -ForegroundColor Green "Folder" $CSVSavePath "already exists`n"
        } 

    Write-Host -ForegroundColor Yellow "`nScript now validating if Excel module installed ..."

    If (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host -ForegroundColor Green "Excel module is currently installed`n"
        } 
    else {
        Write-Host "Excel module needed, installing ..."
        Install-Module ImportExcel -Scope CurrentUser -Force # Install-Module ImportExcel -Force
    }
    pause  
}

Function ListAdminUsers {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
    ForEach-Object { `
    # Prepares File Names for export
    $MailDomain = "@" + $_.MailDomain
    $CSVSavePathAdmins = $CSVSavePath + '\Admins' + $MailDomain + '.csv'
    Write-Host -ForegroundColor Yellow "`n- Login in for listing admin users on Partner Center:" $MailDomain "`n"
    # $ExecutingUser = $SecUser + $MailDomain
    # $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass) 
    Connect-MsolService –Credential $O365Cred
    Write-Host -ForegroundColor Green "Working on it, this might take some time ...`n"
    Get-MsolRole | %{$role = $_.name; Get-MsolRoleMember -RoleObjectId $_.objectid} | select @{Name="Role"; Expression = {$role}}, EmailAddress, DisplayName | export-CSV $CSVSavePathAdmins -NoTypeInformation
    }
#pause
}

Function ListUsers {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
 ForEach-Object { `
$MailDomain = "@" + $_.MailDomain
#$MailDomain = $_.MailDomain.trim()
# Prepares File Names for export
$CSVSavePathUsers = $CSVSavePath + '\Users' + $MailDomain + '.csv'

Write-Host -ForegroundColor Yellow "`n- Login in for listing all users on Partner Center:" $MailDomain "`n"
# $ExecutingUser = $SecUser + $MailDomain
# $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass) 
Connect-MsolService –Credential $O365Cred
$AllUsers = @()
$users = Get-MsolUser -All
Write-Host -ForegroundColor Green "Working on it, this might take some time ...`n"
    foreach ($users_iterator in $users){ 
 
            $user_displayname = $users_iterator.displayname 
            $user_principal_name = $users_iterator.userprincipalname 
            $user_object_id = $users_iterator.objectid 
            $user_type = $users_iterator.UserType 
            $user_role_name = (Get-MsolUserRole -ObjectId $user_object_id).name 
            $user_secgroup_name = (Get-MsolGroup -isAgentRole -UserPrincipalName $user_principal_name).DisplayName
            
        $Properties = @{
		
			DisplayName = $user_displayname
			UserPrincipalName = $user_principal_name 
			UserType = $user_type
			RoleName = $user_role_name
			UserSecGroupName = $user_secgroup_name
            ObjectId = $user_object_id 
		}
        
		$AllUsers += New-Object -Type PSObject -Property $Properties	
    }

    # Multiple Roles assigned RoleName = System.Object[]
    # @{Name='RoleName'; Expression={[string]::join(";", ($_.RoleName))}}
    $AllUsers | select DisplayName, UserPrincipalName, UserType, RoleName, UserSecGroupName, ObjectId
    $AllUsers | select DisplayName, UserPrincipalName, UserType, RoleName, UserSecGroupName, ObjectId | export-CSV $CSVSavePathUsers -NoTypeInformation
    # TRY CATCH -> Write-Host -ForegroundColor Green "Operation Complete .. `n"
}
#pause
}

Function NewUsers {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
 ForEach-Object { `

    $MailDomain = "@" + $_.MailDomain
    # Prepares File Names for export
    $CSVSavePathNewUsers = $CSVSavePath + '\NewUsers' + $MailDomain + '.csv'
    $UsageLocation = $_.UsageLocation
    $Country = $_.Country
    #$CSVUsersOnboarding = $CSVSavePath + '\UsersOnboarding.csv'
    $AddedUsers = @()
    
    Write-Host -ForegroundColor Yellow "`n- Login for adding New Users on Partner Center:" $MailDomain "`n"
    # $ExecutingUser = $SecUser + $MailDomain
    # $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass)    
    # Conection to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        #Import-Csv -Path $Users_Csv |
        Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Users |
            foreach { `
                #Clear pasted spaces
                $FirstName = $_.FirstName.trim()
                $LastName = $_.LastName.trim()
                $AdminGroup = $_.AdminGroup
                $ADRole = $_.RoleName
                $UPN = $_.LoginId + $MailDomain
                $DisplayName = $FirstName + " " + $LastName + " " + $_.Position
                
                # USE FUNCTION UserExists # Add / Improve check user Exists by 7 letters of surname in UPN or LastName from Display Name
                # $LastName = """*" + $_.LastName + "*"""
                # $UserExists = (Get-MsolUser -SearchString $LastName).UserPrincipalName # | Select UserPrincipalName # -UserPrincipalName $UPN -ErrorAction SilentlyContinue
                $UserExists = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue
                                
                If ($UserExists)  {
                Write-Host "User $UPN already exists. Updating Agent role to $AdminGroup"
                $Comments = "User $UPN already existed. Updated Agent role to $AdminGroup"
                    $GroupID = Get-MsolGroup | Where-Object {$_.DisplayName -eq $AdminGroup}
                    # Check user correct memberships
                    # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                    # $UpdatedGroupID = Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                    $CurrentGroup = Get-MsolGroup -isAgentRole -UserPrincipalName $UPN
                    
                    If ($CurrentGroup.ObjectId -ne $GroupID.ObjectID) {
                    Remove-MsoLGroupMember -GroupObjectId $CurrentGroup.ObjectId -GroupmemberObjectId $UserExists.ObjectId
                    $User = Get-MsolUser -UserPrincipalName $UPN
                    Add-MsolGroupMember `
                        -GroupObjectId $GroupID.ObjectId `
                        -GroupMemberType 'User' `
                        -GroupMemberObjectId $User.ObjectId #-ErrorAction Inquire

                    # Add Admin role Get-MsolRole -RoleName 'Company Administrator' | Sort Name | Select Name,Description
                    # Company Administrator         Can manage all aspects of Azure AD and Microsoft services that use Azure AD identities.
                    # Billing Administrator         Can perform common billing related tasks like updating payment information.
                    # User Account Administrator    Can manage all aspects of users and groups, including resetting passwords f

                    Write-Host "and AD role to" $_.RoleName "if changed`n"
                    $Comments = $Comments + "and AD role to $ADRole"
                        If ($ADRole) { #-ne $ActualRole) { # If NOt empty ??
                           Remove-MsolRoleMember -RoleName "Company Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                           Remove-MsolRoleMember -RoleName "Billing Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                           Remove-MsolRoleMember -RoleName "User Account Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                           Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $ADRole #-ErrorAction SilentlyContinue
                        }
                    }
                } Else {
                    # Add Restore User if deleted by cleanup??
                    $NewUser = New-MsolUser `
                        -UserPrincipalName $UPN `
                        -FirstName $FirstName `
                        -LastName $LastName `
                        -DisplayName $DisplayName `
                        -Department 'Customer Operations' `
                        -Title 'Analyst' `
                        -Password 'P@$$w0rd123' `
                        -UserType 'Member' `
                        -StreetAddress 'EMEA' `
                        -Office 'Managed Services EMEA' `
                        -City 'EMEA' `
                        -State 'EMEA' `
                        -UsageLocation $UsageLocation `
                        -Country $Country `
                        -PostalCode '28082' `
                        -PhoneNumber '0844 499 0365' 
                    
                    # Removed for formating output:(Pending to test for results)
                    Write-Host "Creating User $UPN `n"
                    # Review getting accounts to be sent out for credentials
                    #| Export-Csv -Path $PSScriptRoot + '\UsersAccounts.csv' -NoTypeInformation -Append
                    
                    # Obtaining Agent Group ObjectId for Admin agent ("AdminAgents","HelpdeskAgents","SalesAgents")
                    $GroupID = Get-MsolGroup | Where-Object {$_.DisplayName -eq $AdminGroup}

                    # Check users correct memberships
                    # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                    
                    $User = Get-MsolUser -UserPrincipalName $UPN
                    Add-MsolGroupMember `
                        -GroupObjectId $GroupID.ObjectId `
                        -GroupMemberType 'User' `
                        -GroupMemberObjectId $User.ObjectId #-ErrorAction Inquire
                    
                    # Add Admin role Get-MsolRole -RoleName 'Company Administrator' | Sort Name | Select Name,Description
                    # Company Administrator         Can manage all aspects of Azure AD and Microsoft services that use Azure AD identities.
                    # Billing Administrator         Can perform common billing related tasks like updating payment information.
                    # User Account Administrator    Can manage all aspects of users and groups, including resetting passwords for limited admins.

                        If ($ADRole) {
                        Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $ADRole -ErrorAction SilentlyContinue
                        }
                    $Comments = "New user added"
                }

                $Properties = @{
                    Password = 'P@$$w0rd123'
                    UserPrincipalName = $UPN
                    Country = $Country
                    DisplayName = $DisplayName
                    AgentRole = $AdminGroup
			        ADRole = $ADRole
			        Comments = $Comments 
		        }
        
		        $AddedUsers += New-Object -Type PSObject -Property $Properties
            }
    $AddedUsers | select Password, UserPrincipalName, Country, AgentRole, ADRole, Comments | export-CSV $CSVSavePathNewUsers -NoTypeInformation
}
#pause
}

Function PasswordReset {
Try {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
 ForEach-Object { `
    Try { 
    
    $MailDomain = "@" + $_.MailDomain
        
    Write-Host -ForegroundColor Yellow "`n- Login for Password Reset on Partner Center:" $MailDomain
    # $ExecutingUser = $SecUser + $MailDomain
    # $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass)
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Users |
            foreach {
                $UPN = $_.LoginId + $Maildomain
                $NewPassword = Set-MsolUserPassword -UserPrincipalName $UPN -NewPassword 'P@$$w0rd123'
                Write-Host -ForegroundColor Green "`n User Password for $UPN reset to $NewPassword"
                }
    } Catch {
    Write-Host "`n ## Error Capturing Users CSV for Password Reset"
    $_.Exception.Message
    }
  }
} Catch {
Write-Host "`n ## Error Capturing Domain List CSV for Password Reset"
$_.Exception.Message
}
pause
}

Function MFAMethodReset {
Try {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
 ForEach-Object { `
    Try { 
    
    $MailDomain = "@" + $_.MailDomain
        
    Write-Host -ForegroundColor Yellow "`n- Login for MFA method reset on Partner Center:" $Maildomain
    # $ExecutingUser = $SecUser + $MailDomain
    # $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass)
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Users |
            foreach {
                $UPN = $_.LoginId + $Maildomain
                Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $UPN
                Write-Host -ForegroundColor Green "`n MFA methods for user $UPN have been reset"
                }
    } Catch {
    Write-Host -ForegroundColor Red "`n ## Error Capturing Users CSV"
    $_.Exception.Message
    }
  }
} Catch {
Write-Host -ForegroundColor Red "`n ## Error Capturing Domain List CSV"
$_.Exception.Message
}
pause
}

Function DeleteUsers {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
 ForEach-Object { `
    
    #$MailDomain = $_.MailDomain.trim()
    $MailDomain = "@" + $_.MailDomain
    
    Write-Host -ForegroundColor Yellow  "`n- Login for users deletion on Partner Center: $MailDomain `n"
    # $ExecutingUser = $SecUser + $MailDomain
    # $O365Cred = New-Object System.Management.Automation.PSCredential ($ExecutingUser, $SecPass)
    Connect-MsolService –Credential $O365Cred
    
        # Remove users based on csv
        Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Users |
            foreach {
                
                #$LastName = """*" + $_.LastName + "*"""
                $LoginId = $_.LoginId
                $UPN = $LoginId + $MailDomain
                $UserExists = (Get-MsolUser -SearchString $LoginId -ErrorAction SilentlyContinue).UserPrincipalName
                
                If (!$UserExists)  { Write-Host -ForegroundColor Green  "User $LoginId doesn't exist at:" $MailDomain "`n"
                } Else { Get-MsolUser -UserPrincipalName $UPN | Remove-MsolUser -Force;
                      Write-Host -ForegroundColor Green "User $LoginId deleted from:" $MailDomain "`n" 
                      Write-Host -ForegroundColor "Would you lie to permanently DELETE User $LoginId from: " $MailDomain "?`n" 
                      Get-MsolUser -SearchString $LoginId -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin
                      }
                }

    Write-Host -ForegroundColor Yellow "The following users are now on the recycle bin: `n"
    # List deleted users on recycle bin.
    (Get-MsolUser -ReturnDeletedUsers -ErrorAction SilentlyContinue).SignInName
    Write-Host "`n To delete all users from Recycle bin use: `n Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force `n"
    #Get-MsolUser -SearchString $LoginId -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force
 }
 #pause
}

Function Show-Menu {
     param (
           [string]$Title = 'Partner Center Onboarding'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "`n1: Press '1' to Onboard New user/s."
     Write-Host "2: Press '2' to Reset password of existing user/s."
     Write-Host "3: Press '3' to Reset MFA methods of existing user/s"
     Write-Host "4: Press '4' to Offboard Existing user/s."
     Write-Host "5: Press '5' to List All admin privileged accounts."
     Write-Host "6: Press '6' to List All users accounts."
     Write-Host "7: Press '7' to Update Users and Directories Files in use."
     Write-Host "Q: Press 'Q' to quit."
     Write-Host "`nTo confirm your choice, please click Enter"
}

Function Menu{

    do
    {
         Show-Menu
         $input = Read-Host "Please make a selection"
         switch ($input)
         {
               '1' {NewUsers}
               '2' {PasswordReset}
               '3' {MFAMethodReset} 
               '4' {DeleteUsers} 
               '5' {ListAdminUsers} 
               '6' {ListUsers} 
               '7' {UpdateFilesCaptured} 
               'q' {return}
         }
         pause
    }
    until ($input -eq 'q')
}

UpdateFilesCaptured
PrepareFolders
Menu