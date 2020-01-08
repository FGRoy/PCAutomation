## CONNECTING TO PARNER CENTER:
Import-Module MSOnline

## Write-Host "Details for Partner Center Credentials (same password and login needed for all accounts `
## if MFA enabled please allow some time for the script to complete)..."
## * # Remove next three lines comments plus $O365Cred below for user and password prompt for every PC Account # * # Ad1rectory.
#$SecUser = Read-Host -Prompt "Enter your partner center account username (before the @) "
#$SecPass = Read-Host -Prompt "Enter your partner center account password "
#$SecPass = ConvertTo-SecureString $SecPass -AsPlainText -Force 

#Path for csv formated list of users and Partner Center domains. RELATIVE PATH to the dir script is running 
$PCDomains_Csv = Read-Host -Prompt "Enter the location for the csv file containing PC Domains List: (PCDomains_EMEA.csv)" 
$PCDomains_Csv = $PSScriptRoot + '\' + $PCDomains_Csv
$UserstoDelete_Csv = Read-Host -Prompt "Enter the location for the csv file containing Users List: (UserstoDeletePC.csv)"
$UserstoDelete_Csv = $PSScriptRoot + '\' + $UserstoDelete_Csv

#######################################################################################################################
#######################################################################################################################
# Offboard user from all PC -Force (For no prompt) # Partner Center login iterations for connection ###################
#######################################################################################################################

Import-Csv -Path $PCDomains_Csv | 
 ForEach-Object { `
    $FullUser = $SecUser + $_.Maildomain
    $Maildomain = $_.Maildomain
    $UsageLocationPC = $_.UsageLocation

    Write-Host "`n- Login Credentials Used for Partner Center:" $FullUser

    # * # Uncomment next line for user and password NO-prompt # * #
    #$O365Cred = New-Object System.Management.Automation.PSCredential ($FullUser , $SecPass)

    # Conexion to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred
    
        # Remove users based on csv
        Import-Csv -Path $UserstoDelete_Csv |
            foreach {
                $UPN = $_.Firstname + "." + $_.Lastname + $Maildomain
                $UserExists = Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue

                If (!$UserExists)  { Write-Host "User $UPN doesn't exist at:" $Maildomain "`n"
                } Else { Get-MsolUser -UserPrincipalName $UPN | Remove-MsolUser -Force }
                }

    Write-Host "The following users are now on the recycle bin: `n #To delete all users from Recycle: `n Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force `n"
    # List deleted users on recycle bin.
    Get-MsolUser -ReturnDeletedUsers
 }

# Delete all recycle bin users permanetly.
# Get-MsolUser -ReturnDeletedUsers | Remove-MsolUser -RemoveFromRecycleBin -Force
