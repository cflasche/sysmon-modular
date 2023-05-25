# The shared sysmon folder has to be globally accessible from all systems
$shared_sysmon_folder = '\\domain\NETLOGON\Sysmon'

# The local path is used by this script to log and track sysmon
$local_sysmon_folder = 'C:\Program Files\Sysmon'
$max_log_file_size = 10 # This is in KB

# For reporting in Windows Eventlog ensure that the log source is available.
New-EventLog -LogName Application -Source "Sysmon Update"

# Create local path if it does not exist
if (!(Test-Path -PathType Container $local_sysmon_folder)){
    $log_message = "Local folder does not exist - creating it"
    Write-Host $log_message
    Write-EventLog -LogName Application -Source "Sysmon Update" -EntryType Information -EventId 1 -Message $log_message
    New-Item -ItemType Directory -Force -Path $local_sysmon_folder
}

function Add-Log ($message){
    $log_time = Get-Date -Format "yyyy-dd-MMTHH:mm:ssK" 
    "" + $log_time + " " + $message | Out-File -FilePath $log_output_file -Append
    Write-Host $message
    Write-EventLog -LogName Application -Source "Sysmon Update" -EntryType Information -EventId 1 -Message $message
}

function Install-Sysmon {
    # The command below installs Sysmon
    & $exe "-accepteula" "-i" $sysmon_configuration
    # Output the version of Sysmon being installed to output path
    $sysmon_current_version | Out-File -FilePath $output_version_file_path -Force
    # Save the loaded configuration file hash
    $sysmon_configuration_file_hash | Out-File -FilePath $output_configuration_file_hash_path -Force
}

function Remove-Sysmon {
    # The command below uninstalls Sysmon
    & $exe "-accepteula" "-u"
}

function Update-SysmonConfig {
    # The command below updates Sysmon's configuration
    & $exe "-accepteula" "-c" $sysmon_configuration
    # Output the configuration hash used to the output path
    $sysmon_configuration_file_hash | Out-File -FilePath $output_configuration_file_hash_path -Force
}


$log_output_file = $local_sysmon_folder + "\" + "sysmon_output.txt"
$log_output_backup_file = $local_sysmon_folder + "\" + "sysmon_output.old"
if(Test-Path -Path $log_output_file){
    $log_message = "Prior output file found"
    Write-Host $log_message
    Write-EventLog -LogName Application -Source "Sysmon Update" -EntryType Information -EventId 1 -Message $log_message
    $log_size = (Get-Item $log_output_file).length/1KB
    # If log size is greater than or equal to $max_log_file_size
    # rotate the file
    if($log_size -ge $max_log_file_size){
        $log_message = "Rotating log file"
        Write-Host $log_message
        Write-EventLog -LogName Application -Source "Sysmon Update" -EntryType Information -EventId 1 -Message $log_message
        if (Test-Path -Path $log_output_backup_file){
            Remove-Item -Path $log_output_backup_file -Force
            Rename-Item -Path $log_output_file $log_output_backup_file
        } else {
            Rename-Item -Path $log_output_file $log_output_backup_file
        }
    }
}

# Get OS architectrure (32-bit vs 64-bit)
$architecture = $env:PROCESSOR_ARCHITECTURE

# Check if Sysmon is installed
if ($architecture -eq 'AMD64') {
    $service = Get-Service -name Sysmon64 -ErrorAction SilentlyContinue
    $exe = $shared_sysmon_folder + "\" + "Sysmon64.exe"
} else {
    $service = Get-Service -name Sysmon -ErrorAction SilentlyContinue
    $exe = $shared_sysmon_folder + "\" + "Sysmon.exe"
}
if ($null -eq $service) {
    Add-Log "No Sysmon service found."
}

# Reading remote file version information from server.
$sysmon_configuration = $shared_sysmon_folder + "\sysmonconfig.xml"
$sysmon_configuration_file_hash = (Get-FileHash -algorithm SHA1 -Path ($shared_sysmon_folder + "\" + "sysmonconfig.xml")).Hash
$sysmon_current_version = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exe).FileVersion
if ($null -eq $sysmon_configuration_file_hash -Or $null -eq $sysmon_current_version) {
    Add-Log "Unable to determine server file versions. Potentially missing network connection. Trying again next time."
    Return
}

# Reading local file version information 
$output_version_file_path = $local_sysmon_folder + "\" + "sysmon_version.txt"
$output_configuration_file_hash_path = $local_sysmon_folder + "\" + "sysmon_configuration_file_hash.txt"
Add-Log ("Sysmon Version found on server: " + $sysmon_current_version + "; Configuration Hash: " + $sysmon_configuration_file_hash)

# Install Sysmon if it is not installed
if ($null -eq $service) {
    Add-Log "Installing Sysmon"
    Install-Sysmon
    $installed_version = $sysmon_current_version
    $installed_configuration_hash = $sysmon_configuration_file_hash
} else {
    Add-Log "Local Sysmon service found"
    # If Sysmon is installed, get the installed version
    $installed_version = Get-Content -Path $output_version_file_path
    # Also get the current configuration hash
    $installed_configuration_hash = Get-Content -Path $output_configuration_file_hash_path
    Add-Log ("Installed Sysmon version: " + $installed_version + "; Configuration Hash: " + $installed_configuration_hash)
}

# If Sysmon is installed, check if the version needs upgraded
if ($installed_version -ne $sysmon_current_version) {
    Add-Log "Local Sysmon version does not match - Reinstalling"
    Remove-Sysmon
    Install-Sysmon
} else {
    Add-Log "Local Sysmon version matches shared repository version"
    # Check if Sysmon's configuration needs updated
    # Not necessary if Sysmon reinstalled due to version mismatch
    if ($installed_configuration_hash -ne $sysmon_configuration_file_hash){
        Add-Log "Local Sysmon configuration out of sync - Updating"
        Update-SysmonConfig
    } else {
        Add-Log "Local Sysmon configuration matches current configuration"
    }
}
