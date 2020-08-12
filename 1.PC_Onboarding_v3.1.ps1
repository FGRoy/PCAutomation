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
#$global:SecUser = Read-Host -Prompt "Enter your partner center account username (before the @) "
#$global:SecPass = Read-Host -Prompt "Enter your partner center account password "
#$global:SecPass = ConvertTo-SecureString $SecPass -AsPlainText -Force
########################################################################################################################
# ADAPTATIONS: #########################################################################################################
# Option to search in Global Audit or separate Audit functionality #############
# Check PCUsers.xlsx exists on update files captured
# Creates Excluded Admin file -> to be a separated function or manually provided
# ### Filter if file is open close excel ??
########################################################################################################################
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
$global:LogFile = $null

## Lastlogon deletion
<#Function ObtainLastLogon {
############ 365 forensics Install-Module -Name Hawk -Verbose ############
#Obtain User LastLogon 
#########Check Complinace module installed ########
#Autopopulate username from login prompts
#Capture the current time information -30Days
#$LoginAgeLimit = 30 
#$LoginAge = (Get-Date).AddDays(-30)#.ToShortDateString() #-Format "MM/dd/yyyy HH:mm"
$now = Get-Date
$LoginAge = ($now.adddays(-30)).ToString('MM/dd/yyyy HH:mm')
$date = $now.ToString('MM/dd/yyyy HH:mm')
#Sec & Compliance Connection with MFA
Connect-IPPSSession -UserPrincipalName ($UserCredential).UserName
#Sec & Compliance Connection witouth MFA
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
########Check EXOL management Installed ########
Import-Module ExchangeOnlineManagement; Get-Module ExchangeOnlineManagement
#ExchangeOnline V2 Connection with MFA
#$EXOSession=New-ExoPSSession -UserPrincipalName ($UserCredential).UserName
#Import-PSSession $EXOSession -Prefix EXO
Connect-ExchangeOnline -UserPrincipalName ($UserCredential).UserName -ShowProgress $true
#$Audit = Search-UnifiedAuditLog -StartDate $LoginAge -EndDate $date -RecordType AzureActiveDirectoryStsLogon -Operations UserLoggedIn -ResultSize 5000 #|Select UserIds,Operations,CreationDate 
$Audit = Search-UnifiedAuditLog -StartDate $LoginAge -EndDate $date -RecordType AzureActiveDirectoryStsLogon -Operations UserLoggedIn -UserIds ($Admin).EmailAddress   #-ResultSize 5000 
$ConvertAudit = $Audit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
$ConvertAudit | Select-Object UserId,Operation,CreationTime,ClientIP | ft
# -UserIds "fgr@GBPinsightemeacsp.onmicrosoft.com" -ResultSize 5000
# Workload,ObjectID,SiteUrl,SourceFileName,UserAgent
# AuditData.ActorIpAddress #-SessionId "UserLastLogon_byID1"-SessionCommand ReturnNextPreviewPage #UserLoginFailed
}#>
## UpdateFiles no GUI
<# Function UpdateFilesCaptured {
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
    #Write-Host "`n"
    #If (-not(Test-Path $global:PCUsers_Csv)){
    #    Write-Host "The users file does not exist, try again`n"
    #    UpdateFilesCaptured
    #}
}#>


Function UpdateFilesCapturedGUI {
param ( 
    [string]$Title = 'Partner Center Onboarding' 
)
    cls
    Write-Host "================ $Title ================"
    #Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir script is running 
        Write-Host "`nEnter the location for the xlsx file containing Users List (PCUsers.xlsx)"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
        $null = $FileBrowser.ShowDialog()
    $global:PCUsers_Csv = $FileBrowser.FileName
    #Path for storing output documents
    $global:Date = Get-Date -Format "d_MM_yyyy"
    $global:CSVSavePath = $PSScriptRoot + '\' + $Date
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

    Write-Host -ForegroundColor Yellow "`nScript now validating needed PS modules installed ..."

    If (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host -ForegroundColor Green "Excel module is currently installed"
        } 
    else {
        Write-Host "Excel module needed, installing ..."
        Install-Module ImportExcel -Scope CurrentUser -Force # Install-Module ImportExcel -Force
    }

    If (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host -ForegroundColor Green "Exchange Online module is currently installed"
        } 
    else {
        Write-Host "Exchange Online module needed, installing ..."
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force # Install-Module ImportExcel -Force
    }

    <#If (Get-Module -ListAvailable -Name CredentialManager) {
        Write-Host -ForegroundColor Green "CredentialManager module is currently installed"
        } 
    else {
        Write-Host "Excel module needed, installing ..."
        Install-Module CredentialManager -Scope CurrentUser -Force # Install-Module ImportExcel -Force
    }#>

    ######## TRANSCRIPT ###########################################
    $Logtime = (Get-date).ToString('ddMMyy-HHmm')
    If (!$global:LogFile){
        $global:LogFile = $CSVSavePath + '\LogPCAutomate' + $Logtime + '.txt'
        Write-Host "`n"
        Start-Transcript -Path $global:LogFile
    }
    Write-Host "`n"
    pause  
}

Function ListAdminUsers {
#Import-Csv -Path $PCDomains_Csv | 
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
    ForEach-Object { `
    # Prepares File Names for export
    $MailDomain = "@" + $_.MailDomain
    $CSVSavePathAdmins = $CSVSavePath + '\Admins' + $MailDomain + '.csv' # avoid overwrite + (Uniquename) Get-date
    $CSVSavePathAdminsNoLogin = $CSVSavePath + '\AdminsNotLogged' + $MailDomain + '.csv' # avoid overwrite + (Uniquename) Get-date
    #$CSVSavePathAdminsExceptions = $CSVSavePath + '\AdminsExcluded' + $MailDomain + '.csv' # avoid overwrite + (Uniquename) Get-date
    ############# GLobal Audit Param ###########################
    #$CSVSavePath90Audit = $CSVSavePath + '\AuditPC' + $Logtime + $MailDomain + '.csv'
    #$CSVSavePathRawAudit = $CSVSavePath + '\AuditPC_RAW' + $Logtime + $MailDomain + '.csv'
    Write-Host -ForegroundColor Yellow "`n- Login in for listing admin users on Partner Center:" $MailDomain "`n"
    Try {
    Connect-MsolService –Credential ($O365Cred).Username -ea stop #REview Failed login action 
    Write-Host -ForegroundColor Green "Working on it, this might take some time ...`n"
    $AllPCAdmins = @()
    $AllPCAdmins = Get-MsolRole | %{$role = $_.name; $step++; Write-Progress -Activity "Obtaining Admin Users Info" -Status "$role" -PercentComplete ($step / $role.count * 100);
    Get-MsolRoleMember -RoleObjectId $_.objectid; Write-Progress -Activity "Obtaining Admin Users Info" -Completed; $step=0} | select @{Name="Role"; Expression = {$role}}, EmailAddress, DisplayName 
    ### Filter if file is open close excel ??
    $AllPCAdmins | export-CSV $CSVSavePathAdmins -NoTypeInformation 
    #$PCAdminsExceptions = @()  # oRIGINAL FOR INCLUSION IN FILE
    $CleanUp = Read-Host -Prompt "Would you like to clean up existing login credentials [y/n]"
    Switch ($CleanUp) 
     { 
       Y {### Clean-up 30 days
        $now = Get-Date
        $LoginAge = ($now.adddays(-30)).ToString('MM/dd/yyyy HH:mm')
        $date = $now.ToString('MM/dd/yyyy HH:mm')
        $PCAdminsDel = @()
        ## loADING EXCEPTIONS
        $PCAdminsExceptions = @()  
        Import-Excel -Path $PCUsers_Csv -WorkSheetname Exceptions |
                    foreach-Object { `
                            If ($_.LoginId){
                            $ExemptUser = $_.LoginID.Split('@')
                            $ExemptUser = $ExemptUser[0] + $MailDomain
                            $PCAdminsExceptions += $ExemptUser.tolower() 
                            #$ExemptUser # Debugging
                            }
                    }
        Write-host -ForegroundColor Yellow "Yes: Cleaning Up admin accounts with last login time older than $LoginAge`n" 
        Write-host -ForegroundColor Green "Connecting to Security and Compliance module ...`n" 
        #Sec & Compliance Connection with MFA
        Connect-IPPSSession -UserPrincipalName ($O365Cred).UserName
        #Sec & Compliance Connection without MFA
        #$o365Cred = Get-Credential
        #$SessionSec = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $o365Cred -Authentication Basic -AllowRedirection
        #Import-PSSession $SessionSec -DisableNameChecking -Prefix SEC
        
        Write-host -ForegroundColor Green "`nConnecting to O365 Exchange Management ..." 
        #Exchange Online Connection witouth with MFA
        Connect-ExchangeOnline -UserPrincipalName ($O365Cred).UserName -ShowProgress $true
        #Exchange Online Connection witouth without MFA
        #$SessionEx = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365Cred -Authentication Basic -AllowRedirection
        #Import-PSSession $SessionEx -DisableNameChecking -Prefix EXO
       
        ##### Global Audit 30 Days #################### Around 5 minutes to execute ###############
        ###########################################################################################
        <#$GlobalAudit = Search-UnifiedAuditLog -StartDate $LoginAge -EndDate $date -ResultSize 5000 
        $GlobalAudit | export-CSV $CSVSavePathRawAudit -NoTypeInformation 
        $ConvertGlobalAudit = $GLobalAudit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
        $ConvertGlobalAudit | export-CSV $CSVSavePath90Audit -NoTypeInformation #>
            foreach ($Admin in $AllPCAdmins){
            If (($Admin).EmailAddress){
              If (-not($PCAdminsExceptions.contains(($Admin).EmailAddress.tolower()))) {
              ###### Option to search in Global Audit or separate Audit functionality #############
                $Audit = Search-UnifiedAuditLog -StartDate $LoginAge -EndDate $date -RecordType AzureActiveDirectoryStsLogon -Operations UserLoggedIn -UserIds ($Admin).EmailAddress #-ResultSize 5000 
                $ConvertAudit = $Audit | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
                If ($ConvertAudit.count -eq 0) {
                    Write-host "Deleting admin user account ... `n $Admin"
                    #Get-MsolUser -UserPrincipalName ($Admin).EmailAddress | Remove-MsolUser
                    $PCAdminsDel += $Admin
                } # ENd If not found in audit
              } Else {
                Write-host "EXEMPT ... `n " ($Admin).EmailAddress
              } # ENd If COntained in Exempt
            } else {Write-host "EXEMPT API ... `n"}# ENd If clear API admins with no email field
            } # ENd Foreach admin not loged-in
        ### FOr MFA Sessions removal EXOL ###
        Get-PSSession | Remove-PSSession #-Session $SessionEXO / $SessionSEC
        $PCAdminsDel | export-CSV $CSVSavePathAdminsNoLogin -NoTypeInformation 
        Write-host -ForegroundColor Green "No: Cleanup process completed, admins file exported at`n $CSVSavePathAdmins"
        } # End Selection Y
       Default {Write-host -ForegroundColor Green "No: Process completed, admins file exported at`n $CSVSavePathAdmins"} 
     } # ENd clean UP switch 
    } Catch { # [System.Management.Automation.RuntimeException] {
        #Write-Host -ForegroundColor Red "Failed to connect to Partner Center $_.MailDomain`n"
        Write-warning $Error[0]
    } # End Connecction Try-Catch
    } # ENd for-each PC
#pause
} # End Function List Admins

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
            $step++
            $user_displayname = $users_iterator.displayname 
            $user_principal_name = $users_iterator.userprincipalname 
            $user_object_id = $users_iterator.objectid 
            $user_type = $users_iterator.UserType 
            Write-Progress -Activity "Obtaining Users Info" -Status "$user_displayname" -PercentComplete ($step / $Users.count * 100)
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
    Write-Progress -Activity "Obtaining Users Info" -Completed; $step=0
    # Multiple Roles assigned RoleName = System.Object[]
    # @{Name='RoleName'; Expression={[string]::join(";", ($_.RoleName))}}
    ## Display exported Users info
    # $AllUsers | select DisplayName, UserPrincipalName, UserType, RoleName, UserSecGroupName, ObjectId
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
  Try{
    # Conection to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred -ea stop
    
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
                # Initialise Comments
                $Comments = ""
                # $UserExists = (Get-MsolUser -SearchString $LastName).UserPrincipalName # | Select UserPrincipalName # -UserPrincipalName $UPN -ErrorAction SilentlyContinue
                $UserExists = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue
                                
                If ($UserExists)  {
                    Write-Host "User $UPN already exists."
                    $Comments = "User $UPN already existed."
                    # Logic for update / remove role if empty
                    #If ($AdminGroup) { 
                        $GroupID = Get-MsolGroup | Where-Object {$_.DisplayName -eq $AdminGroup}
                        # Check user correct memberships
                        # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                        # $UpdatedGroupID = Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                        $CurrentGroup = Get-MsolGroup -isAgentRole -UserPrincipalName $UPN
                        If ($CurrentGroup.ObjectId -ne $GroupID.ObjectID) {
                            Write-Host "Updating Agent role to $AdminGroup"
                            $Comments = "Updated Agent role to $AdminGroup"
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
                        } # End IF not Eq -> Update Role
                    #} else { # If AdminGroup empty Remove existing Role
                        #Remove-MsoLGroupMember -GroupObjectId $CurrentGroup.ObjectId -GroupmemberObjectId $UserExists.ObjectId
                    #} # End IF AdminGroup
                    If ($ADRole) { # If NOt empty # -ne $ActualRole) {
                        Write-Host "Updating AD role to" $_.RoleName "if changed`n"
                        $Comments = $Comments + "and AD role to $ADRole"
                        Remove-MsolRoleMember -RoleName "Company Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                        Remove-MsolRoleMember -RoleName "Billing Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                        Remove-MsolRoleMember -RoleName "User Account Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                        Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $ADRole #-ErrorAction SilentlyContinue
                    #} else { # If ADRole empty Remove existing Role
                        #Remove-MsolRoleMember -RoleName "Company Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                        #Remove-MsolRoleMember -RoleName "Billing Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                        #Remove-MsolRoleMember -RoleName "User Account Administrator" -RoleMemberEmailAddress $UPN -ErrorAction SilentlyContinue
                    } # End IF ADRole
                } Else { # User does not exit Add
                    # Generate Random Passwd
                    $length = 14 ## characters
                    $nonAlphaChars = 3
                    $Password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
                    # Generate Random Passwd
                    # $minLength = 5 ## characters
                    # $maxLength = 10 ## characters
                    # $length = Get-Random -Minimum $minLength -Maximum $maxLength
                    # $nonAlphaChars = 5
                    # $password = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)
                    # Add Restore User if deleted by cleanup??
                    $NewUser = New-MsolUser `
                        -UserPrincipalName $UPN `
                        -FirstName $FirstName `
                        -LastName $LastName `
                        -DisplayName $DisplayName `
                        -Department 'Customer Operations' `
                        -Title 'Analyst' `
                        -Password $Password `
                        -UserType 'Member' `
                        -StreetAddress 'EMEA' `
                        -Office 'Managed Services EMEA' `
                        -City 'EMEA' `
                        -State 'EMEA' `
                        -UsageLocation $UsageLocation `
                        -Country $Country `
                        -PostalCode '28082' `
                        -PhoneNumber '0844 499 0365' 
                    
                    If ($NewUser) { # If user created
                        # Removed for formating output:(Pending to test for results)
                        Write-Host "Creating User $UPN `n"
                        # Review getting accounts to be sent out for credentials
                        #| Export-Csv -Path $PSScriptRoot + '\UsersAccounts.csv' -NoTypeInformation -Append
                    
                        # Obtaining Agent Group ObjectId for Admin agent ("AdminAgents","HelpdeskAgents","SalesAgents")
                        $GroupID = Get-MsolGroup | Where-Object {$_.DisplayName -eq $AdminGroup} -ErrorAction SilentlyContinue

                        # Check users correct memberships
                        # Get-MsolGroupMember -GroupObjectId $GroupID.ObjectId
                        If($GroupID){ # IF $Admin GRoup Provided on Excel Add to New User
                            $User = Get-MsolUser -UserPrincipalName $UPN
                            Add-MsolGroupMember `
                                -GroupObjectId $GroupID.ObjectId `
                                -GroupMemberType 'User' `
                                -GroupMemberObjectId $User.ObjectId #-ErrorAction Inquire
                        } else {
                            Write-Host "No agent permisions defined for user: $UPN `n"
                        } # End else GroupID
                    
                        # Add Admin role Get-MsolRole -RoleName 'Company Administrator' | Sort Name | Select Name,Description
                        # Company Administrator         Can manage all aspects of Azure AD and Microsoft services that use Azure AD identities.
                        # Billing Administrator         Can perform common billing related tasks like updating payment information.
                        # User Account Administrator    Can manage all aspects of users and groups, including resetting passwords for limited admins.

                        If ($ADRole) {
                            Add-MsolRoleMember -RoleMemberEmailAddress $UPN -RoleName $ADRole -ErrorAction SilentlyContinue
                        } else {
                            Write-Host "No AD organization permisions defined for user: $UPN `n"
                        } # End else ADRole
                        Write-Host "User added $UPN `n"
                        $Comments = "New user added"
                    } else { # En If NewUser created
                        Write-Host "No user account has been created for user: $UPN `n"
                    } # End else NewUser Creation failed
                } # End else NewUser Add

                $Properties = @{
                    Password = $Password # 'P@$$w0rd123'
                    UserPrincipalName = $UPN
                    Country = $Country
                    DisplayName = $DisplayName
                    AgentRole = $AdminGroup
			        ADRole = $ADRole
			        Comments = $Comments 
		        } # End Object AddedUser Properties
        
		        $AddedUsers += New-Object -Type PSObject -Property $Properties
            } # End for each User in Excel
        if ($AddedUsers.count -eq 0)    {
	        Write-Host -ForegroundColor Red "No Users were added or updated for $MailDomain"
        } else {
        $AddedUsers | select Password, UserPrincipalName, Country, AgentRole, ADRole, Comments | export-CSV $CSVSavePathNewUsers -NoTypeInformation
        } # ENd If added users eq 0
   } Catch { # [System.Management.Automation.RuntimeException] {
        Write-Host -ForegroundColor Yellow "Failed to connect to Partner Center $_.MailDomain"
   } # End Connecction Try-Catch
 } # End for each PC Domain in Excel
    #pause
} # End Function Add New USer

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
                # Set Random Password ###############################################################################################
                # Set-MsolUserPassword -UserPrincipalName $UPN -ForceChangePassword
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
                      Write-Host -ForegroundColor Yellow "Would you like to permanently DELETE User $LoginId from: " $MailDomain "?`n" 
                      ### Update message from pop-up for confirmation ### Disable for automation
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

Function PCLastLogonDate {
Import-Excel -Path $PCUsers_Csv -WorkSheetname PC_Domains |
    ForEach-Object { `
    # Prepares File Names for export
    $MailDomain = "@" + $_.MailDomain
    Write-Host -ForegroundColor yellow "`n- Login in for keeping Partner Center access:" $MailDomain "`n"
    # $ExecutingUser = $SecUser + $MailDomain
    Try{
        Connect-MsolService –Credential $O365Cred -ea Stop
        Write-Host -ForegroundColor Green "`n Time count restored ...`n"
    }Catch{
        Write-host -Foregroundcolor Magenta "`n Login failed ...." 
    } # End Try If Credentials already exist
    }
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
     Write-Host "8: Press '8' to Update PC Last Logon Date."
     Write-Host "Q: Press 'Q' to quit."
     Write-Host "`nTo confirm your choice, please click Enter"
}

Function Menu{

    do
    {
         Show-Menu
         $menuinput = Read-Host "Please make a selection"
         switch ($menuinput)
         {
               '1' {NewUsers}
               '2' {PasswordReset}
               '3' {MFAMethodReset} 
               '4' {DeleteUsers} 
               '5' {ListAdminUsers} 
               '6' {ListUsers} 
               '7' {UpdateFilesCapturedGUI} 
               '8' {PCLastLogonDate} 
               'q' {return}
         }
         pause
    }
    until ($input -eq 'q')
}

UpdateFilesCapturedGUI
PrepareFolders
Menu
Write-Host "`n"
Stop-Transcript