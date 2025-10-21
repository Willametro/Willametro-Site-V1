# AD Computer Windows Update Scanner

A robust PowerShell application with GUI to scan Active Directory computers for Windows Update information.

## Features

- **GUI Interface**: Modern WPF-based graphical interface
- **OU Browser**: Interactive tree view to browse and select Active Directory Organizational Units
- **Flexible Scanning**: Choose to include or exclude sub-OUs
- **Concurrent Scanning**: Configurable parallel scanning (1-50 concurrent scans)
- **Comprehensive Information**:
  - Computer online/offline status
  - Operating System version
  - Last boot time
  - Number of installed updates (from history)
  - Number of available updates
  - Pending reboot status
- **Progress Tracking**: Real-time progress bar and status updates
- **Export Capability**: Export results to CSV for further analysis
- **Robust Error Handling**: Gracefully handles offline computers and errors
- **Stop Functionality**: Ability to stop long-running scans

## Requirements

### Software Requirements
- Windows PowerShell 5.1 or PowerShell 7+
- Active Directory PowerShell Module (RSAT)
- .NET Framework 4.5 or higher

### Permissions Required
- Active Directory read permissions
- WinRM enabled on target computers
- Administrative access to target computers (for Windows Update queries)
- Firewall rules allowing:
  - ICMP (ping)
  - WMI (TCP 135, 445)
  - WinRM (TCP 5985 for HTTP, 5986 for HTTPS)

### Installing RSAT (Active Directory Module)

**Windows 10/11:**
```powershell
# Using Settings app
Settings > Apps > Optional Features > Add a feature > RSAT: Active Directory Domain Services

# Or using PowerShell (as Administrator)
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

**Windows Server:**
```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

## Usage

### Running the Application

1. **Launch PowerShell as Administrator** (recommended for best results)

2. **Set execution policy** if needed:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Run the script**:
   ```powershell
   .\AD-WindowsUpdate-Scanner.ps1
   ```

### Using the GUI

1. **Select Organizational Unit**:
   - Click "Browse OU..." button
   - Navigate the tree to find your desired OU
   - Click "OK" to select

2. **Configure Scan Options**:
   - Check/uncheck "Include Sub-OUs" as needed
   - Set "Max Concurrent Scans" (default: 10)
     - Lower values: Slower but less network/CPU impact
     - Higher values: Faster but more resource intensive

3. **Start Scanning**:
   - Click "Start Scan"
   - Monitor progress in the status section
   - Use "Stop Scan" to abort if needed

4. **View Results**:
   - Results appear in real-time in the grid
   - Sortable by clicking column headers
   - Columns:
     - **Computer Name**: Name of the computer
     - **Status**: Online/Offline/Error
     - **OS Version**: Operating system information
     - **Last Boot**: Last system boot time
     - **Installed Updates**: Count of updates in history
     - **Available Updates**: Count of pending updates
     - **Pending Reboot**: Yes/No/Unknown
     - **Error**: Any error messages

5. **Export Results**:
   - Click "Export to CSV"
   - Choose save location
   - Results saved in CSV format

## Troubleshooting

### Common Issues

**"Failed to connect to Active Directory"**
- Ensure you're on a domain-joined computer
- Verify RSAT is installed
- Check you have AD read permissions

**"Unable to ping computer"**
- Computer may be offline
- Firewall may be blocking ICMP
- Network connectivity issues

**"Update Query: Access Denied"**
- You need administrative permissions on target computers
- Check that your account has appropriate rights

**"WinRM errors"**
- WinRM may not be enabled on target computers
- Enable with: `Enable-PSRemoting -Force`
- Check GPO settings for WinRM configuration

**Slow scanning**
- Reduce "Max Concurrent Scans"
- Many computers may be offline (timeout delays)
- Network latency issues

### Performance Tips

1. **Start with smaller OUs** to test before scanning large environments
2. **Adjust concurrent scans** based on your network capacity
3. **Schedule scans** during off-peak hours for large environments
4. **Filter offline computers** from results after first scan

## Technical Details

### How It Works

1. **AD Query**: Retrieves computer objects from specified OU using Active Directory module
2. **Connectivity Test**: Pings each computer to check if online
3. **WMI Query**: Gets OS information and last boot time via WMI
4. **Registry Check**: Checks pending reboot status via registry keys
5. **Windows Update Query**: Uses COM object (Microsoft.Update.Session) to query update history and available updates
6. **Parallel Processing**: Uses PowerShell jobs for concurrent scanning

### Update Detection Methods

**Installed Updates Count**:
- Queries Windows Update history (last 1000 entries)
- Counts successful installation operations

**Available Updates Count**:
- Searches for updates where:
  - IsInstalled = False
  - Type = 'Software'
  - IsHidden = False

**Pending Reboot Detection**:
- Component Based Servicing registry key
- Windows Update Auto Update registry key
- PendingFileRenameOperations registry value

## Security Considerations

- Script requires administrative credentials for full functionality
- Uses PowerShell Remoting (WinRM) - ensure proper security configuration
- Does not store credentials
- Read-only operations on target systems
- No modifications made to target computers

## Limitations

- Requires Windows environment with Active Directory
- Target computers must be Windows-based
- Requires network connectivity to target computers
- Performance depends on network speed and computer count
- Some updates may not be detected if Windows Update service has issues

## License

Created for administrative use in Active Directory environments.

## Version History

- **v1.0** - Initial release
  - OU browser
  - Update scanning
  - CSV export
  - Concurrent scanning
  - Progress tracking
