Write-Host "-----------------------------------"
Write-Host "|   *******'s USER TERM SCRIPT    |"
Write-Host "-----------------------------------"
Write-Host "This script is based on 'xxx.docx' located in EIT documentation. Also be advised this script is unforgiving of errors. You will be prompted to login, use admin acct where possible"
[System.Net.WebRequest]::DefaultWebProxy.Credentials =
[System.Net.CredentialCache]::DefaultCredentials

Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted

Write-Host "Loading..."
#Install-Module MSOnline
Write-host "."
#Install-Module AzureADPreview -Force
Write-host ".."
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement
Write-Host "Modules loaded"

#Query admin for some info
$Admin = "admin" #Read-Host -Prompt "Please enter your admin username" 
$FN = Read-Host -Prompt "Enter the FIRSTNAME of the employee you wish to terminate (e.g. 'John')"
$LN = Read-Host -Prompt "Enter the LASTNAME of the employee you wish to terminate (e.g. 'Smith')"
$User = Read-Host -Prompt "Enter the USERNAME of the employee you wish to terminate (e.g. 'jsmith')" 
$Name = $FN + " " + $LN
$Email = $User + "@suffix"

#DN function
Function Get-DistinguishedName ($User) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$User))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
$strDN = Get-DistinguishedName $User 

#Remove all memberships in "Member of" tab.
$principalGroups = (Get-ADPrincipalGroupMembership $user | where{$_.sAMAccountName -ne "Domain Users"}).sAMAccountName
    Foreach($principalGroup in $principalGroups){
				Remove-ADPrincipalGroupMembership $user -MemberOf $principalGroup -Confirm:$false}
Start-Sleep -Seconds 10

#Open Active Directory Users and Computers and move the user to (domain.contoso.local\Exchange\o365disabledusers. #Move-ADObject -Identity $strDN 
Get-ADUser -Identity $User | Move-ADObject -TargetPath "OU=O365DisabledUsers,OU=Exchange,DC=domain,DC=contoso,DC=local" #-TargetServer "domaindc1pw.domain.contoso.local"
      
#Get the new DN
#Function Get-DistinguishedName ($User) 
#{  
#   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
#   $searcher.Filter = "(&(objectClass=User)(samAccountName=$User))" 
#   $result = $searcher.FindOne() 
# 
#   Return $result.GetDirectoryEntry().DistinguishedName 
#} 
#Gets the new DN
$strDN = Get-DistinguishedName $User 

#Clear the "Company" field.
Get-ADUser $User | Set-ADUser -clear company

#Change password
$password = (ConvertTo-SecureString -string N0tEmployed!321 -AsPlainText -Force)
Set-ADAccountPassword -Identity $User -Reset -NewPassword $password

#Mailbox stuff
Write-Host "Waiting 60s wait to let the AD replication happen"
Start-Sleep -s 60

Write-Host "Disregard the 'remoteMailbox.RemoteRecipientType must include ProvisionMailbox', BUT verify in Exchange that the conversion to Shared happened. That means the user was migrated to 365 and it doesn't like how we're converting to shared mailbox. See https://support.microsoft.com/en-us/help/4515271/can-t-convert-a-migrated-remote-user-mailbox-to-shared-in-exchange-ser for more info."

Write-Host "Disregard but patiently wait out any errors that mention a Watson dump."

#https://portal.office.com/adminportal/home#/homepage, Exchange>Mailboxes, search for username. "Convert to Shared Mailbox".
Connect-ExchangeOnline -ShowBanner:$false
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mail1pw.domain.contoso.local/powershell -Authentication Kerberos -AllowRedirection #-Credential $UserCredential 
Import-PSSession $Session -DisableNameChecking -AllowClobber
Set-RemoteMailbox -Identity $User -Type Shared -HiddenFromAddressListsEnabled $true #Changed to below line to test
#Set-Mailbox -Identity $User -Type Shared -HiddenFromAddressListsEnabled $true 

#-ConnectionUri http://exchange1pw.domain.contoso.local/powershell

#Creating reformatted date string
$Date = Get-Date -Format FileDate
$yr = $Date.Substring(0,4)
$mo = $Date.Substring(4,2)
$da = $Date.Substring(6,2)
$Date = $mo + $da + $yr

#Remove 365 licenses. Doc specifies noting which type of license is being removed; if/else with append output of username and license type to file.
Connect-MsolService
$file = "c:\users\$Admin\desktop\$User$Date.txt"
Add-Content $file "USER: `n$User"

#License notation block
$cmd = (Get-MsolUser -UserPrincipalName $Email).Licenses.ServiceStatus
If ($cmd | Out-String -Stream | Select-String -Pattern "EXCHANGE_S_STANDARD_") {
    Write-Host "User has G1"
    $Sku = "G1"
}
Elseif ($cmd | Out-String -Stream | Select-String -Pattern "EXCHANGE_S_ENTERPRISE_") {
	Write-Host "User has G3"
    $Sku = "G3"
}
Else {
		Write-Host "User has some other license, probably F3. PLEASE VERIFY IT MANUALLY AFTER THE SCRIPT COMPLETES."
        $Sku = "UNKNOWN"
}
Add-Content $file "LICENSE: `n$Sku"

#License removal
#Set-MsolUserLicense -UserPrincipalName $User �RemoveLicenses $Lic
(get-MsolUser -UserPrincipalName $Email).licenses.AccountSkuId |
foreach{
    Set-MsolUserLicense -UserPrincipalName $Email -RemoveLicenses $_
}

Write-host "Per xxx's email from 4/30/21, mailbox delegation is to be done via EAC. The mailbox delegation block has therefore been disabled."

#If access to MB is needed, delegate access to that user, Full Access. User can access the MB via webmail.suffix and selecting "Open other mailbox".
#$Delegation = Read-Host -prompt "Will this mailbox need to be delegated to anyone? Y/N"
#$Delegation = $Delegation.ToLower()
#If ($Delegation -eq "y") {
#    While ($Delegation -eq "y") {
#        $Remainer = Read-Host -prompt "Please enter the delegate's EMAIL address"
#        Add-Content $file "Mailbox delegate: `n$Remainer"
#        Add-Mailboxpermission �Identity $Email �User $Remainer �Accessrights FullAccess #Added 3/15/21, below line was commented out
#        #Set-MailboxFolderPermission -Identity '$Email':\inbox -User $Remainer -AccessRights Editor #was Set-MailboxPermission -identity $Email -User $Remainer -AccessRights Editor, doesn't work
#        $Delegation = Read-Host -prompt "Will this mailbox need to be delegated to anyone? Y/N"
#        $Delegation = $Delegation.ToLower()
#        }
#        }

#Closes previous EOL session as it is not needed
Remove-PSSession $Session

$homedirpath = "\\domainhome\home\$user"
$homedocspath = "\\domainhome\home\$User\My Documents"
#if/else to check if user has a home folder, and if so, to execute various actions upon it
$there = test-path $homedirpath

#if/else to check if user has a home folder, and if so, to execute various actions upon it
$docsthere = test-path $homedocspath

if ($there -eq $True) {
    write-host "User has an H:\ drive folder"
    #new folder ownership block
    $ACL = Get-ACL "\\domainhome\home\$user"#added quotes 12/7, may need to change to $homedirpath to address ctor issue
    $Owner = New-Object System.Security.Principal.NTAccount($Admin)
    $ACL.SetOwner($Owner)
    Set-Acl -Path \\domainhome\home\$user -AclObject $ACL

    $src = "\\domainhome\home\$User"
    $dst = "\\domainhome\separations\$User" #added $User 1/4/21
    $docs = "\\domainhome\home\$User\My Documents"

    #take ownership of my documents
    if ($docsthere -eq $True) {
        $ACL = Get-ACL -Path "\\domainhome\home\$User\My Documents" #changed to path from $docs 2/12 #$docs #1/27 added -Path
        $Owner = New-Object System.Security.Principal.NTAccount($Admin)
        $ACL.SetOwner($Owner)
        Set-Acl -Path $docs -AclObject $ACL #'\\domainhome\home\$user\My Documents' -AclObject $ACL 
        }
        else{}

    #Get full control of My Documents so we can move the data
    $principal = "domain\$admin"
    $FileSystemAccessRights=[System.Security.AccessControl.FileSystemRights]"FullControl"
    $InheritanceFlags=[System.Security.AccessControl.InheritanceFlags]�ContainerInherit, ObjectInherit�
    $PropagationFlags=[System.Security.AccessControl.PropagationFlags]�None�
    $AccessControl=[System.Security.AccessControl.AccessControlType]�Allow�
    $permission = @($principal), @($FileSystemAccessRights), @($InheritanceFlags), @($PropagationFlags), @($AccessControl)
    if ($docsthere -eq $True) {
        $acl = Get-Acl $docs
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule -ArgumentList $permission
        $acl.SetAccessRule($AccessRule)
        $acl | Set-Acl $docs
        }
        else{}

    takeown /f "\\domainhome\home\$User" /r /d Y

    #Move the home folder to \\domainhome\separation, append the name with the date of separation, e.g. jdoe01012021
    $src = "\\domainhome\home\$User"
    $dst = "\\domainhome\separations\$User$Date" #added $User 1/4/21
    Move-Item -Path $src -Destination $dst -Force
    #Rename-Item "\\domainhome\separations\$User" "\\domainhome\separations\$User$Date"
    #WILL ADD 'DELETE ORIGINAL' LATER ONCE I SEE GOOD TESTING

    #If access to home folder needed, assign r-- to the users who need access. Append output of contact person, directory name, and current date + 30days to a different file for you to check later to confirm the directory can be deleted.
    #while loop while ans = yes, assign r-- to specified user. n breaks out

    $homedir = "\\domainhome\separations\$User$Date"  #quotes correct 1/7/21
    $ans = Read-Host -Prompt "Will anyone need access to the departing user's home folder? Y/N"
    $ans = $ans.ToLower()
    While ($ans -eq "y") {
        $Delegate = Read-Host -Prompt "Please enter the USERNAME of the user to get read-only access"
        Add-Content $file "File delegate: `n$Delegate"
        #Grant-FileShareAccess -Name $homedir -AccessRight "Read" -AccountName domain\$Delegate
        $ACL = Get-ACL -Path $homedir
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Delegate,"Read","Allow")
        $ACL.SetAccessRule($AccessRule)
        $ACL | Set-Acl -Path $homedir
        (Get-ACL -Path $homedir).Access | Format-Table IdentityReference,FileSystemRights,AccessControlType,IsInherited,InheritanceFlags -AutoSize
        $ans = Read-Host -Prompt "Will anyone need access to the departing user's home folder? Y/N"
        $ans = $ans.ToLower()
        }
    }
    else {
    write-host "User does not have an H:\ drive folder"
    }

#Append (Disabled Account) to the account name
$strDN = Get-DistinguishedName $User
$disName = $Name + " (Disabled Account)"
Rename-ADObject -Identity $strDN -NewName $disName
Disable-ADAccount -Identity $User

Write-Host "User has been renamed to " $disName

Write-Host "Don't forget to follow up in 30 days to verify the folder can be deleted. A record of this transaction is on your desktop."