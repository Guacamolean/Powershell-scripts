IF (!(Test-Path C:\Logs)){
    New-item "C:\Logs" -ItemType Directory
}

$logTime = Get-Date -Format "hh.mm tt"
$logDate = Get-Date -Format "MMM dd"
$logFile = "C:\Logs\QA Check.log"
$bios = Get-WmiObject -Class Win32_Bios

#------------FUNCTIONS----------

function Write-Log 
{
    $output = $args[0] | Out-String
    Add-Content -Path $logFile -Value $output.Replace("`n","")
}

function Get-WindowsStatus 
{
    $osVersion = Get-WmiObject Win32_OperatingSystem | Select-Object Caption
    $windowsStatus = Get-WmiObject SoftwareLicensingProduct -Filter "ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f'" | Where-Object {$_.LicenseStatus -eq "1" -or $_.LicenseStatus -eq "2"}
    Switch ($windowsStatus.LicenseStatus){
	    "1" {Write-Log "$($osVersion.Caption) : Activated."}
	    "2" {Write-Log "$($osVersion.Caption) : Not activated."}
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
            Write-Log "$officeVersion : Activated."
        }
        ELSE {
            Write-Log "$officeVersion : Not activated."
        }
    }
    ELSE {
        Write-Log "Office : Not installed."
    }
}

function Get-Kaseya 
{
    $kaseya = Get-Service -DisplayName "Kaseya*"
    $kaseyaDirectory = Get-ChildItem 'C:\Program Files (x86)' -recurse -include kaseyad.ini | Select-Object DirectoryName
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
        Write-Log "Kaseya checking into $($install.Group) on $($install.Server)."
    }
    }
    ELSE {
        Write-Log "Could not find Kaseya installed!"
    }
}

function Get-Trend 
{
    $trend = Get-Service -DisplayName "Trend Micro*Security Agent"
    IF ($trend){
        Write-Log "Trend Micro : Installed."
        Write-Log "$($trend.DisplayName) : $($trend.Status)"
    }
	ELSE {
	    Write-Log "Trend Micro : Not installed."
    }
}

function Get-FlashPlayer ($Version)
{
    $flashPlayer = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | `
    Get-ItemProperty | Where-Object {$_.DisplayName -like "Adobe Flash Player*$version"}
  
    IF ($flashPlayer){
        Write-Log "Adobe Flash Player $version $($flashPlayer.DisplayVersion) : Installed."
    }
    else {
        Write-Log "Adobe Flash Player $version : Not installed."
    }
}

function Get-Java
{
    $java = Get-ChildItem -Path  HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall,HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | `
    Get-ItemProperty | Where-Object {$_.DisplayName -like "Java ? Update *"}
    
    IF ($java){
        foreach ($javaInstall in $java){
        Write-Log "$($javaInstall.DisplayName) : Installed."
        } 
    }
    Else {
        Write-Log "Java : Not installed."
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
        Write-Log "$($wiredDNS.Interface) DNS : $($wiredDNS.Configuration)."}
    IF ($wirelessDNS){
        Write-Log "$($wirelessDNS.Interface) DNS : $($wirelessDNS.Configuration)."}
}

function Get-Drivers
{
    $drivers = Get-WmiObject Win32_PNPEntity | Where-Object {$_.ConfigManagerErrorCode -ne 0} | Select-Object Name
    IF ($drivers){
        foreach ($driver in $drivers){
            Write-Log "$($driver.Name) is reporting a driver error."}
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
        Write-Log "$driveLetter $percentFree% available."
    }
}

function Get-SystemRestore
{
    $systemRestore = Get-WmiObject -Namespace 'root\default' -Class SystemRestoreConfig
    IF ($systemRestore.RPSessionInterval -eq "1"){
        Write-Log "System Restore : Enabled using $($systemRestore.DiskPercent)% disk space."
    }
    ELSE {
        Write-Log "System Restore : Not enabled."
    }
}

#----------EXECUTION----------

Write-Log "********************************QA Check************************************************************"
Write-Log "Software check started at $logTime on $logDate."
Write-Log "Hostname : $env:COMPUTERNAME."
Write-Log "Serial Number : $($bios.SerialNumber)"
Get-WindowsStatus
Get-OfficeStatus
Get-Kaseya
Get-Trend
Get-FlashPlayer -Version "ActiveX"
Get-FlashPlayer -Version "PPAPI"
Get-FlashPlayer -Version "NPAPI"
Get-Java
Get-DnsSettings
Get-Drivers
Get-DiskInfo
Get-SystemRestore
Write-Log "****************************************************************************************************"
Write-Log ""
