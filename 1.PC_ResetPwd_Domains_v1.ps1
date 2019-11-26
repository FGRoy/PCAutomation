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
$UserstoReset_Csv = Read-Host -Prompt "Enter the location for the csv file containing Users List: (UserstoResetPC.csv)"
$UserstoReset_Csv = $PSScriptRoot + '\' + $UserstoReset_Csv

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
    #$UsageLocationPC = $_.UsageLocation
    
    ## * # Un-Comment next two lines for user and password no prompt for every PC Account # * #
    Write-Host "`n- Login Credentials Used for Partner Center:" $FullUser
    #$O365Cred = New-Object System.Management.Automation.PSCredential ($FullUser , $SecPass)
    
    # Conexion to Office 365 Partner Center Account
    Connect-MsolService –Credential $O365Cred
    
        # Add users based on csv
        Import-Csv -Path $UserstoReset_Csv |
            foreach {
                
                $UPN = $_.Firstname + "." + $_.Lastname + $Maildomain
                #$DisplayName = $_.Firstname + " " + $_.Lastname + " " + $_.Position #" [EMEA IMS L1]"
                #$User = Get-MsolUser -UserPrincipalName $UPN
                #Set-ADAccountPassword -Identity '' -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "p@ssw0rd" -Force)
                Set-MsolUserPassword -UserPrincipalName $UPN -NewPassword 'P@$$w0rd123'
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