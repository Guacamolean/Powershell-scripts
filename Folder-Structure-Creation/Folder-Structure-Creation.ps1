<#
Uses an external csv file to read in information.  Must contain 5 headers.
    Folder  -   name of folder to be created with drive letter
    Read    -   any users that should have Read access, seperated by (;) semicolon
    Modify  -   any users that should have Modify access, seperated by (;) semicolon
    Full    -   any users that should have Full Control access, seperated by (;) semicolon
    Domain  -   name of domain users will be a part of, insert on first line only
#>

Function Set-FullAccessRule($Share, $Domain, $User)
{
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$domain\$user","FullControl","ContainerInherit, ObjectInherit","None","Allow")
    $acl = Get-Acl $share
    $acl.SetAccessRule($accessRule)
    Set-Acl $share $acl
}

Function Set-ModifyAccessRule($Share, $Domain, $User)
{
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$domain\$user","Modify","ContainerInherit, ObjectInherit","None","Allow")
    $acl = Get-Acl $share
    $acl.SetAccessRule($accessRule)
    Set-Acl $share $acl
}

Function Set-ReadAccessRule($Share, $Domain, $User)
{
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("$domain\$user","ReadAndExecute","ContainerInherit, ObjectInherit","None","Allow")
    $acl = Get-Acl $share
    $acl.SetAccessRule($accessRule)
    Set-Acl $share $acl
}

# Creates variable using imported information from csv file
$csv = Import-Csv shared.csv

$domain = $csv[0].domain

foreach($folder in $csv){
    $share = $folder.Folder
    IF (!(Test-Path $share)){
        New-Item -ItemType Directory -Path $share 
        IF ($folder.Full){
            $users = $folder.Modify.Split(';') | ForEach-Object {$_.Trim()}
            foreach($user in $users){
                Set-FullAccessRule -Share $share -Domain $domain -User $user                
            }
        }        
        IF ($folder.Modify){
            $users = $folder.Modify.Split(';') | ForEach-Object {$_.Trim()}
            foreach($user in $users){
                Set-ModifyAccessRule -Share $share -Domain $domain -User $user                
            }
        }
        IF ($folder.Read){
            $users = $folder.Read.Split(';') | ForEach-Object {$_.Trim()}
            foreach($user in $users){
                Set-ReadAccessRule -Share $share -Domain $domain -User $user
            }
        }
    }
}
