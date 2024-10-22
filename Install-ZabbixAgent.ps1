# Define variables
$downloadUrl = "https://cdn.zabbix.com/zabbix/binaries/stable/7.0/7.0.4/zabbix_agent-7.0.4-windows-amd64-openssl.msi"  # Update with the actual latest Zabbix Agent URL
$tempPath = "C:\Temp"
$msiPath = "$tempPath\zabbix_agent.msi"
$zabbixServer = "<zabbixServer,zabbixServer>"  # Update with your Zabbix server address
$zabbixActiveServer = "zabbixActiveServer;zabbixActiveServer"  # Update with your Zabbix active server address

# Stop and uninstall the existing Zabbix Agent (if present)
Write-Host "Stopping existing Zabbix Agent (if running)..."
Stop-Service -Name "Zabbix Agent" -ErrorAction SilentlyContinue

Write-Host "Uninstalling existing Zabbix Agent..."
& "C:\Program Files\Zabbix Agent\zabbix_agentd.exe" --uninstall
Get-Service | Where-Object { $_.Name -eq "Zabbix Agent" }

# Create temp directory if it doesn't exist
if (-Not (Test-Path -Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath
}

# Download the latest Zabbix agent MSI
Write-Host "Downloading Zabbix Agent..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $msiPath

# Install the Zabbix agent
Write-Host "Installing Zabbix Agent..."
Start-Process msiexec.exe -ArgumentList "/i", $msiPath, "/quiet", "/norestart" -Wait

# Configure Zabbix Agent settings
Write-Host "Configuring Zabbix Agent settings..."
$zabbixConfPath = "C:\Program Files\Zabbix Agent\zabbix_agentd.conf"  # Path to Zabbix agent configuration file

# Update server and active server in the Zabbix configuration file
(Get-Content $zabbixConfPath) -replace '^(Server=).*', "`$1$zabbixServer" |
    Set-Content $zabbixConfPath

(Get-Content $zabbixConfPath) -replace '^(ServerActive=).*', "`$1$zabbixActiveServer" |
    Set-Content $zabbixConfPath

# Restart the Zabbix agent service to apply the new settings
Write-Host "Restarting Zabbix Agent service..."
try {
    Restart-Service -Name "Zabbix Agent" -ErrorAction Stop
    Write-Host "Zabbix Agent service restarted successfully."
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}

# Cleanup: Remove the downloaded file
Remove-Item $msiPath

Write-Host "Download and installation process completed."
