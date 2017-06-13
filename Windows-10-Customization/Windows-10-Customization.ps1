Function Set-RegistryValues($Path,$Name)
{
    IF (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue){
        Set-ItemProperty -Path $path -Name $name -Value 0
    }
    ELSE {
        New-ItemProperty -Path $path -Name $name -Value 0 -Property DWORD
    }
}

$desktopPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
$myComputer = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
$userFiles = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
$network = "{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}"
$trayPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Explorer"
$trayKey = "EnableAutoTray"

Set-RegistryValues -Path $desktopPath -Name $myComputer
Set-RegistryValues -Path $desktopPath -Name $userFiles
Set-RegistryValues -Path $desktopPath -Name $network
Set-RegistryValues -Path $trayPath -Name $trayKey