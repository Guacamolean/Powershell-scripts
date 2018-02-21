#Function to set "PasswordNeverExpires" flag to true or false for local Windows accounts.  Tested on Windows 10/Server 2012 R2.

#Import the .NET assembly needed for this operation.
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

#Function accepts 2 parameters: $user is the user you want to change the "PasswordNeverExpires" value on, 
#and $Expiry is the boolean value you want to set it to ($true or $false)

Function Set-PasswordExpiry($User, $Expiry) {
    $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Machine
    $userPrincipal = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($contextType, $user)
    $userPrincipal.PasswordNeverExpires = $expiry
    $userPrincipal.Save()
    }
