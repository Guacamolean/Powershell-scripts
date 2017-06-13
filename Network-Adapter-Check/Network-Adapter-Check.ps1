$logFile = "C:\Logs\QA Check.log"

function Write-Log {
    $output = $args[0] | Out-String
    Add-Content -Path $logFile -Value $output.Replace("`n","")
}

IF (!(Test-Path C:\Logs)){
    New-item "C:\Logs" -ItemType Directory
}

function Get-DnsSettings{
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

    $wiredDNS = $results | Where-Object {($_.Interface -eq "Local Area Connection") -or ($_.Interface -eq "Ethernet")}
    $wirelessDNS = $results | Where-Object {$_.Interface -like "Wireless*" -or $_.Interface -like "Wi-Fi"}
}

Write-Log "**********************************Network Adapter Check*********************************************"

$ethernet = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object {($_.Name -like "*Ethernet Connection*") -or ($_.Name -like "*Gigabit*")}
IF ($ethernet.NetEnabled -eq $false){
    $ethernetEnable = $ethernet.enable()
    Switch ($ethernetEnable.ReturnValue){
        "0" {Write-Log "$($ethernet.Name) was found disabled, it has now been successfully re-enabled."}
        "5" {Write-Log "$($ethernet.Name) was found disabled. Received an access denied error when trying to re-enable.  Please enable manually."}
    }
}

Start-Sleep 10
Get-DnsSettings
IF ($wirelessDNS.configuration -eq "Static"){
    netsh.exe interface ipv4 set dns "$($wirelessDNS.Interface)" dhcp
}
IF ($wiredDNS.configuration -eq "Static"){
    netsh.exe interface ipv4 set dns "$($wiredDNS.Interface)" dhcp
}

Write-Log "****************************************************************************************************"
Write-Log ""