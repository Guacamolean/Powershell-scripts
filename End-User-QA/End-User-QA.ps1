IF (!(Test-Path C:\Logs)){
    New-item "C:\Logs" -ItemType Directory
}

$logTime = Get-Date -Format "hh.mm tt"
$logDate = Get-Date -Format "MMM dd"
$logFile = "C:\Logs\QA Check.log"
$bios = Get-WmiObject -Class Win32_Bios

#------------FUNCTIONS----------

function Write-Log {
    Param(
        [Parameter(
            Mandatory=$True
        )]
        [String]
        $String,

        [Parameter(
            Mandatory=$True
        )]
        [String]
        $Color
    )
    Add-Content -Path $logFile -Value $String
    Write-Host $String -ForegroundColor $Color
}

function Get-WindowsStatus 
{
    $osVersion = Get-WmiObject Win32_OperatingSystem | Select-Object Caption
    $windowsStatus = Get-WmiObject SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" | Where-Object {$_.LicenseStatus -eq "1" -or $_.LicenseStatus -eq "2"}
    Switch ($windowsStatus.LicenseStatus){
	    "1" {Write-Log "$($osVersion.Caption) : Activated." "Green"}
	    "2" {Write-Log "$($osVersion.Caption) : Not activated." "Red"}
    }
}

function Get-OfficeStatus 
{
    IF (Test-Path 'C:\Program Files (x86)\Microsoft Office\Office14\ospp.vbs'){
        $officePath = "C:\Program Files (x86)\Microsoft Office\Office14"
        $officeVersion = "Office 2010"
        } 
        ELSEIF (Test-Path 'C:\Program Files (x86)\Microsoft Office\Office15\ospp.vbs'){
            $officePath = "C:\Program Files (x86)\Microsoft Office\Office15"
            $officeVersion = "Office 2013"
        }
        ELSEIF (Test-Path 'C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs'){
            $officePath = "C:\Program Files (x86)\Microsoft Office\Office16"
            $officeVersion = "Office 2016"
        }
        ELSEIF (Test-Path 'C:\Program Files\Microsoft Office\Office16\ospp.vbs'){
            $officePath = "C:\Program Files\Microsoft Office\Office16"
            $officeVersion = "Office 2016"
        }
        ELSE {
        $officePath = $Null
    }
    IF ($officePath){
        $officeStatus = cscript $officePath\ospp.vbs /dstatus | Select-String -Pattern "License Status" -SimpleMatch
        IF ($officeStatus -like "*LICENSED*"){
            Write-Log "$officeVersion : Activated." "Green"
        }
        ELSE {
            Write-Log "$officeVersion : Not activated." "Yellow"
        }
    }
    ELSE {
        Write-Log "Office : Not installed." "Red"
    }
}

function Get-Kaseya 
{
    $kaseya = Get-Service -DisplayName "Kaseya*"
    $kaseyaDirectory = Get-ChildItem 'C:\Program Files (x86)' -recurse -include kaseyad.ini -ErrorAction SilentlyContinue | Select-Object DirectoryName
    IF ($kaseya){
        $kaseyaObject = @(
            ForEach ($directory in $kaseyaDirectory){
            $kaseyaSettings = Get-Content -Path "$($directory.DirectoryName)\kaseyad.ini"
            $kaseyaGroup = $kaseyaSettings | Select-String -Pattern "User_Name"
            $kaseyaGroup = $kaseyaGroup -Split "\s+" | Select-Object -Last 1 
            $kaseyaServer = $kaseyaSettings | Select-String -Pattern "^Server_Name"
            New-Object pscustomobject -Property @{
                'Group' = $kaseyaGroup -replace "$env:COMPUTERNAME.",""
                'Server' = $kaseyaServer -Split "\s+" | Select-Object -Last 1
                }	
            }
            )
    ForEach ($install in $kaseyaObject){
        Write-Log "Kaseya checking into $($install.Group) on $($install.Server)." "Green"
    }
    }
    ELSE {
        Write-Log "Could not find Kaseya installed!" "Red"
    }
}

function Get-Trend 
{
    $trend = Get-Service -DisplayName "Trend Micro*Security Agent"
    IF ($trend){
        Write-Host "Trend Micro : Installed." -ForegroundColor "Green"
        Switch($trend.Status){
            'Running'{Write-Log "$($trend.DisplayName) : $($trend.Status)" "Green"}
            'Stopped'{Write-Host "$($trend.DisplayName) : $($trend.Status)" "Yellow"}
        }
    }
	ELSE {
	    Write-Log "Trend Micro : Not installed." "Red"
    }
}

function Get-FlashPlayer ($Version)
{
    $flashPlayer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | `
    Get-ItemProperty | Where-Object {$_.DisplayName -like "Adobe Flash Player*$version"}
  
    IF ($flashPlayer){
        Write-Log "Adobe Flash Player $version $($flashPlayer.DisplayVersion) : Installed." "Green"
    }
    else {
        Write-Log "Adobe Flash Player $version : Not installed." "Yellow"
    }
}

function Get-Java
{
    $java = Get-ChildItem -Path  HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | `
    Get-ItemProperty | Where-Object {$_.DisplayName -like "Java ? Update *"}
    
    IF ($java){
        foreach ($javaInstall in $java){
        Write-Log "$($javaInstall.DisplayName) : Installed." "Green"
        } 
    }
    Else {
        Write-Log "Java : Not installed." "Red"
    }
}

function Get-DnsSettings
{
    $dnsSettings = netsh interface ipv4 show dns | Where-Object {![string]::isnullorempty($_)} | ForEach-Object {$_.trim()}
    $hash = @{}
    $ip = New-Object System.Collections.ArrayList
    $results = @(
        foreach ($line in $dnsSettings){
            switch -Regex ($line){
                '^configuration' {
                    if ($hash.Keys.Count -gt 0){
                        New-Object pscustomobject -Property $hash
                        $hash = @{}
                    }
                    $hash.Add('Interface', $line.split('"')[1])
                    continue
                }
                '(static).*servers\:\s+([\d.]+|none)' {
                    $hash.Add('Configuration', $Matches[1])
                    $ip.Add($Matches[2]) | Out-Null
                    continue
                }
                'dns.*(dhcp):\s+([\d.]+|none)' {
                    $hash.Add('Configuration', $Matches[1])
                    $ip.Add($Matches[2]) | Out-Null
                    continue
                }
                '([\d.]+)' {
                    $ip.Add($Matches[1]) | Out-Null
                    continue
                }
                'register' {
                    if ($ip) {
                        $hash.Add('IP', [string]::Join('; ', $ip.ToArray()))
                        $ip = New-Object System.Collections.ArrayList
                    }
                    $hash.Add('Suffix', $line.split(':')[1].trim())
                    continue
                }
                default {
                    Write-Warning "Don't know what to do with this: $line"
                }
            }
        }
        New-Object pscustomobject -Property $hash
    )
    $wiredDNS = $results | Where-Object {$_.Interface -like "Local*"}
    $wirelessDNS = $results | Where-Object {$_.Interface -like "Wireless*" -or $_.Interface -like "Wi-Fi"}

    IF ($wiredDNS){
        Switch ($wiredDNS.Configuration){
            'DHCP' {Write-Log "$($wiredDNS.Interface) DNS : $($wiredDNS.Configuration)." "Green"}
            'Static' {Write-Log "$($wiredDNS.Interface) DNS: $($wiredDNS.Configuration)." "Yellow"}
        }
    }
    IF ($wirelessDNS){
        Switch ($wirelessDNS.Configuration){
            'DHCP' {Write-Log "$($wirelessDNS.Interface) DNS : $($wirelessDNS.Configuration)." "Green"}
            'Static' {Write-Log "$($wirelessDNS.Interface) DNS : $($wirelessDNS.Configuration)." "Yellow"}
        }
    }
}

function Get-Drivers
{
    $drivers = Get-WmiObject Win32_PNPEntity | Where-Object {$_.ConfigManagerErrorCode -ne 0} | Select-Object Name
    IF ($drivers){
        foreach ($driver in $drivers){
            Write-Log "$($driver.Name) is reporting a driver error." "Red"} 
    }
    Else {
        Write-Log "No driver errors being reported." "Green"
    }
}

function Get-DiskInfo
{
    $disks = Get-WmiObject -Class Win32_LogicalDisk | Where-Object {$_.DriveType -eq "3"}
    foreach ($disk in $disks){
        $freeSpace = $disk.FreeSpace/1GB
        $size = $disk.Size/1GB
        $percentFree = [math]::Round($freeSpace/$size*100)
        $driveLetter = $disk.DeviceID
        IF ($percentFree -lt 25){
            Write-Log "$driveLetter $percentFree% available." "Red"
        }
        else {
            Write-Log "$driveLetter $percentFree% available" "Green"
        }
    }
}   

function Get-SystemRestore
{
    $systemRestore = Get-WmiObject -Namespace 'root\default' -Class SystemRestoreConfig
    IF ($systemRestore.RPSessionInterval -eq "1"){
        Write-Log "System Restore : Enabled using $($systemRestore.DiskPercent)% disk space." "Green"
    }
    ELSE {
        Write-Log "System Restore : Not enabled." "Red"
    }
}

#----------EXECUTION----------

Write-Log "********************************QA Check************************************************************" "White"
Write-Log "Software check started at $logTime on $logDate." "White"
Write-Log "Hostname : $env:COMPUTERNAME." "White"
Write-Log "Serial Number : $($bios.SerialNumber)" "White"
Write-Log "Checking Windows status...." "White"
Get-WindowsStatus 
Write-host "Checking Office status..."
Get-OfficeStatus
Write-Host "Checking Kaseya status..."
Get-Kaseya
Write-Host "Checking Trend status..."
Get-Trend
Write-Host "Checking for Flash Player installs..."
Get-FlashPlayer -Version "ActiveX"
Get-FlashPlayer -Version "PPAPI"
Get-FlashPlayer -Version "NPAPI"
Write-Host "Checking for Java installs.."
Get-Java
Write-Host "Checking DNS Settings..."
Get-DnsSettings
Write-Host "Checking for driver issues..."
Get-Drivers
Write-Host "Getting disk info..."
Get-DiskInfo
Write-Host "Checking System Restore settings..."
Get-SystemRestore
Write-Log "****************************************************************************************************" "White"
Read-Host "Script complete.  Results can also be found in $logFile - Press Enter to exit..."