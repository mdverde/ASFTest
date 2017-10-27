#######################################################################
#
#
#  Script Name: DTCSetup
#  Author: Andrés Durán Hewitt
#  Version: 0.1.0
#  Creation Date: 9/14/2016
#  Last Updated: 12/13/2016
#  Last Updated By: Andrés Durán Hewitt
#
#
#######################################################################


#********* Enable Local MSDTC in the cloud service being run *********#

# Detect OS version
$osVersion = (Get-WmiObject -class Win32_OperatingSystem).Caption
$expectedOSVersion = "2016"

# Setup registry values
$COMPortsRegistryPath = "HKLM:\SOFTWARE\Microsoft\Rpc\Internet"

[hashtable] $COMRegistryItems = @{ "Ports" = "5000-5200"; "PortsInternetAvailable" = "Y"; "UseInternetPorts" = "Y" }

If ($osVersion -match $expectedOSVersion)
{
    Write-Host "Your current version of Windows is $osVersion. Running built-in cmdlets..."
	Write-Host "Checking if network access to MSDTC is already setup..."
	$dtcNetworkSettings = Get-DtcNetworkSetting -DtcName Local
	If ($dtcNetworkSettings.AuthenticationLevel -ne "NoAuth" -or $dtcNetworkSettings.InboundTransactionsEnabled -ne $true -or $dtcNetworkSettings.OutboundTransactionsEnabled -ne $true -or $dtcNetworkSettings.RemoteClientAccessEnabled -ne $true)
    {
	    Set-DtcNetworkSetting –DtcName Local –AuthenticationLevel NoAuth –InboundTransactionsEnabled 1 –OutboundTransactionsEnabled 1 –RemoteClientAccessEnabled 1 –confirm: $false
		Write-Host "Network access to MSDTC has been enabled. Restarting MSDTC..."
		Restart-Service MSDTC
		Write-Host "MSDTC restarted successfully. Setting up firewall rules..."
    }
	Else
	{
		Write-Host "Network access to MSDTC has already been enabled for this virtual machine. Setting up firewall rules..."
	}
    Write-Host "A- Inbound rules"
    # DTC ONLY RULES
    If ((Get-NetFirewallRule -Name MSDTC-RPCSS-In-TCP -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Group "Distributed Transaction Coordinator" -Enabled True -Action Allow -Direction Inbound -Name "MSDTC-RPCSS-In-TCP" -DisplayName "Distributed Transaction Coordinator (RPC-EPMAP)" -Description "Inbound rule for the RPCSS service to allow RPC/TCP traffic for the Kernel Transaction Resource Manager for Distributed Transaction Coordinator service." -Program "%SystemRoot%\system32\svchost.exe" -Protocol TCP -LocalPort RPCEPMap -RemotePort Any -LocalAddress Any -RemoteAddress Any -Profile Any -InterfaceType Any -EdgeTraversalPolicy Block -Service "RPCSS"
    }
    If ((Get-NetFirewallRule -Name MSDTC-KTMRM-In-TCP -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Group "Distributed Transaction Coordinator" -Enabled True -Action Allow -Direction Inbound -Name "MSDTC-KTMRM-In-TCP" -DisplayName "Distributed Transaction Coordinator (RPC)" -Description "Inbound rule for the Kernel Transaction Resource Manager for Distributed Transaction Coordinator service to be remotely managed via RPC/TCP." -Program "%SystemRoot%\system32\svchost.exe" -Protocol TCP -LocalPort RPC -RemotePort Any -LocalAddress Any -RemoteAddress Any -Profile Any -InterfaceType Any -EdgeTraversalPolicy Block -Service "ktmrm"
    }
    If ((Get-NetFirewallRule -Name MSDTC-In-TCP -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Group "Distributed Transaction Coordinator" -Enabled True -Action Allow -Direction Inbound -Name "MSDTC-In-TCP" -DisplayName "Distributed Transaction Coordinator (TCP-In)" -Description "Inbound rule to allow traffic for the Distributed Transaction Coordinator. [TCP]" -Program "%SystemRoot%\system32\msdtc.exe" -Protocol TCP -LocalPort Any -RemotePort Any -LocalAddress Any -RemoteAddress Any -Profile Any -InterfaceType Any -EdgeTraversalPolicy Block
    }
    # DYNAMIC RPC PORT OPENING
    If ((Get-NetFirewallRule -Name TITANIUM-MSDTC-SUPPORT-In-DYN-PORTS -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Enabled True -Action Allow -Direction Inbound -Name "TITANIUM-MSDTC-SUPPORT-In-DYN-PORTS" -DisplayName "Titanium - Dynamic COM Port Opening For RPC Communication [DTC Support]" -Description "Inbound rule to open the configured port range to allow MSDTC to receive incoming requests." -Protocol TCP -LocalPort 5000-5200 -Profile Any -EdgeTraversalPolicy Block
    }
    If ((Get-NetFirewallRule -Name TITANIUM-MSDTC-SUPPORT-In-EPMP-PORT -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Enabled True -Action Allow -Direction Inbound -Name "TITANIUM-MSDTC-SUPPORT-In-EPMP-PORT" -DisplayName "Titanium - Endpoint Mapper Port Opening For RPC Communication [DTC Support]" -Description "Inbound rule to open the configured port in COM Internet Services to allow MSDTC to dynamically allocate a port for a given incoming request." -Protocol TCP -LocalPort 135 -Profile Any -EdgeTraversalPolicy Block
    }
    Write-Host "B- Outbound rules"
    # DTC ONLY RULES
    If ((Get-NetFirewallRule -Name MSDTC-Out-TCP -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Group "Distributed Transaction Coordinator" -Enabled True -Action Allow -Direction Outbound -Name "MSDTC-Out-TCP" -DisplayName "Distributed Transaction Coordinator (TCP-Out)" -Description "Outbound rule to allow traffic for the Distributed Transaction Coordinator. [TCP]" -Program "%SystemRoot%\system32\msdtc.exe" -Protocol TCP -LocalPort Any -RemotePort Any -LocalAddress Any -RemoteAddress Any -Profile Any -InterfaceType Any
    }
    # DYNAMIC RPC PORT OPENING
    If ((Get-NetFirewallRule -Name TITANIUM-MSDTC-SUPPORT-Out-DYN-PORTS -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Enabled True -Action Allow -Direction Outbound -Name "TITANIUM-MSDTC-SUPPORT-Out-DYN-PORTS" -DisplayName "Titanium - Dynamic COM Port Opening For RPC Communication [DTC Support]" -Description "Outbound rule to open the configured port range to allow MSDTC to send outcoming responses." -Protocol TCP -RemotePort 5000-5200 -Profile Any
    }
    If ((Get-NetFirewallRule -Name TITANIUM-MSDTC-SUPPORT-Out-EPMP-PORT -ErrorAction SilentlyContinue) -eq $null)
    {
        New-NetFirewallRule -Enabled True -Action Allow -Direction Outbound -Name "TITANIUM-MSDTC-SUPPORT-Out-EPMP-PORT" -DisplayName "Titanium - Endpoint Mapper Port Opening For RPC Communication [DTC Support]" -Description "Outbound rule to open the configured port in COM Internet Services to allow MSDTC to dynamically allocate a port when responding to a given incoming request." -Protocol TCP -RemotePort 135 -Profile Any
    }
    Write-Host "Firewall rules setup process has finished. Restarting firewall service..."
    Restart-Service MpsSvc -Force
    Write-Host "Firewall restarted successfully. Enabling DCOM Port Range..."
    $dcomRangeEnabled = $true
    If (!(Test-Path -Path $COMPortsRegistryPath))
    {
        New-Item -Path $COMPortsRegistryPath
        $dcomRangeEnabled = $false
    }
    ForEach ($registryItem in $COMRegistryItems.GetEnumerator())
    {
        $keyType = $null
        If ($registryItem.Value.Length -eq 1) { $keyType = "String" } Else { $keyType = "MultiString" }
        If (!(Get-ItemProperty -Path $COMPortsRegistryPath -Name $registryItem.Key -ErrorAction SilentlyContinue))
        {
            New-ItemProperty -Path $COMPortsRegistryPath -Name $registryItem.Key -Value $registryItem.Value -PropertyType $keyType
            $dcomRangeEnabled = $false
        }
    }
    If (!$dcomRangeEnabled)
    {
        Restart-Computer
    }
}
Else
{
    Write-Host "The current version of Windows is not currently supported by this script."
}


## netsh interface ip del wins name="Ethernet 2" all

## netsh interface ip add wins name="Ethernet 2" addr=172.18.19.70 index=0

Write-Host "All taken care of. Signing off... ;-)"