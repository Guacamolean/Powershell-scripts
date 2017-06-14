$computerSystem = Get-WmiObject -Class Win32_ComputerSystem | Select-Object Model
$arguments = @(
    "/i"
    "DisplayLink_Win7-10TH2.msi"
    "/norestart"
    "/quiet"
)

IF ($computerSystem.Model -eq "Latitude 5480" -or $computerSystem.Model -eq "Latitude 7280"){
    Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow
} 