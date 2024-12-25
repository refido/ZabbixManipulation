from config import Config
from pyzabbix import ZabbixAPI
import pandas as pd
import os
from datetime import datetime

# connect
config = Config()
zapi = ZabbixAPI(config.ZabbixAPI)
zapi.login(config.ZabbixUser, config.ZabbixPass)

# Print Zabbix version information
print("\nZabbix Version Information:")
print("-" * 50)
print(f"API Version: {zapi.api_version()}")
print(f"Zabbix Server Version: {zapi.version}")
print("-" * 50 + "\n")

# Define output directory
output_dir = "zabbix_reports"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Define item keys for different OS types
windows_keys = {
    'system.hostname': 'OS Hostname',
    'system.uname': 'OS Description',           # Works on both Windows and Linux
    'system.cpu.load': 'CPU Usage',            # Basic CPU load for Windows
    'vm.memory.size[total]': 'Total Memory',
    'vm.memory.size[free]': 'Free Memory',
    'vfs.fs.size[C:,free]': 'Disk C: Free',    # Windows specific disk monitoring
    'system.uptime': 'System Uptime'
}

# Get all hosts with extended information for debugging
hosts = zapi.host.get(
    selectInterfaces=['interfaceid', 'ip', 'type'],
    selectItems=['itemid', 'key_', 'lastvalue', 'state', 'status', 'error'],
    output=['hostid', 'name', 'status', 'available']
)

# Process each host
data = []
for host in hosts:
    print(f"\nProcessing host: {host['name']}")
    print(f"Host availability status: {host['available']}")  # 0 - unknown, 1 - available, 2 - unavailable
    
    host_data = {'Hostname': host['name']}
    
    # Get IP address
    if 'interfaces' in host:
        for interface in host['interfaces']:
            if interface['type'] == '1':  # 1 is for agent interface
                host_data['IP Address'] = interface['ip']
                print(f"Agent interface IP: {interface['ip']}")
                break
    
    # Debug: Print all available items for this host
    print(f"\nAvailable items for {host['name']}:")
    if 'items' in host:
        for item in host['items']:
            print(f"Item key: {item['key_']}, Value: {item.get('lastvalue', 'N/A')}, "
                  f"State: {item.get('state', 'N/A')}, Error: {item.get('error', 'None')}")
    
    # Collect metrics
    for key, metric_name in windows_keys.items():
        value = 'N/A'
        if 'items' in host:
            for item in host['items']:
                if item['key_'] == key:
                    value = item.get('lastvalue', 'N/A')
                    if value != 'N/A':
                        # Format values based on metric type
                        if 'cpu' in key.lower():
                            try:
                                value = f"{float(value):.2f}%"
                            except (ValueError, TypeError):
                                value = 'N/A'
                        elif 'memory' in key.lower() or 'size' in key.lower():
                            try:
                                value_bytes = float(value)
                                value = f"{value_bytes / (1024**3):.2f} GB"
                            except (ValueError, TypeError):
                                value = 'N/A'
                        elif 'uptime' in key.lower():
                            try:
                                days = float(value) / 86400
                                value = f"{days:.2f} days"
                            except (ValueError, TypeError):
                                value = 'N/A'
        host_data[metric_name] = value
    
    data.append(host_data)

# Create dataframe
df = pd.DataFrame(data)

# Generate timestamp for filename
timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
excel_file = os.path.join(output_dir, f'zabbix_system_metrics_{timestamp}.xlsx')

# Export to Excel
df.to_excel(excel_file, index=False)

print(f"\nData exported to {excel_file}")