$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$loggingFunctionsPath = Join-Path -Path $scriptPath -ChildPath "..\Logging_Function\Logging_Functions.ps1"
. $loggingFunctionsPath
$ScriptVersion = "1.0"
function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$ModuleNames
    )
    # Start logging
    Start-Log -ScriptVersion $ScriptVersion -ScriptPath $PSCommandPath
    # Get the log file path from the start-log function

    foreach ($ModuleName in $ModuleNames) {
        if (Get-Module -ListAvailable -Name $ModuleName) {
            if (-not (Get-Module -Name $ModuleName)) {
                Import-Module $ModuleName
                if (-not (Get-Module -Name $ModuleName)) {
                    $message = "`tFailed to import module '$ModuleName'."
                    Write-Log -Message $message -Level "Error" -sFullPath $global:sFullPath
                } else {
                    $message = "`tModule '$ModuleName' imported successfully."
                    Write-Log -Message $message -Level "OK" -sFullPath $global:sFullPath
                }
            } else {
                $message = "`tModule '$ModuleName' is already imported."
                Write-Log -Message $message -Level "Info" -sFullPath $global:sFullPath
            }
        } else {
            $message = "`tModule '$ModuleName' does not exist."
            Write-Log -Message $message -Level "Error" -sFullPath $global:sFullPath
        }
    }
}
# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.850', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility'
# Define CSV file name
$csvFileName = "Appliances_List.csv"
# Create the full path to the CSV file
$csvFilePath = Join-Path -Path $scriptPath -ChildPath $csvFileName
# Define the path to the credential folder and file
$credentialFolder = Join-Path -Path $scriptPath -ChildPath "credential"
$credentialFile = Join-Path -Path $credentialFolder -ChildPath "credential.txt"
try {
    # Attempt to connect to the appliance
    Connect-OVMgmt -Hostname $ApplianceFQDN -Credential $Credential

    # If the connection is successful, log a success message
    $message = "Successfully connected to : $ApplianceFQDN"
    Write-Log -Message $message -Level "OK" -sFullPath $global:sFullPath
} catch {
    # If a connection already exists, log a message and continue
    if ($_.Exception.Message -like "*already connected*") {
        $message = "Already connected to : $ApplianceFQDN"
        Write-Log -Message $message -Level "Info" -sFullPath $global:sFullPath
    } else {
        # If the connection fails for any other reason, log an error message
        $message = "Failed to connect to : $ApplianceFQDN. Error details: $($_.Exception.Message)"
        Write-Log -Message $message -Level "Error" -sFullPath $global:sFullPath
    }    
}
# Check if the credential folder exists, if not, create it
if (!(Test-Path -Path $credentialFolder)) {
    Write-Log -Message "The credential folder $credentialFolder does not exist. Create it now..." -Level "Warning" -sFullPath $global:sFullPath
    New-Item -ItemType Directory -Path $credentialFolder | Out-Null
    Write-Log -Message "The credential folder $credentialFolder has been created successfully." -Level "OK" -sFullPath $global:sFullPath
}

# If the credential file exists, try to load the credential from it
if (Test-Path -Path $credentialFile) {
    try {
        $credential = Import-Clixml -Path $credentialFile
    } catch {
        Write-Host "Error loading credential file. Please enter your credentials."
        $credential = Get-Credential -Message "Enter your username and password"
        $credential | Export-Clixml -Path $credentialFile
    }
} else {
    # If the credential file doesn't exist, ask for the username and password and store them in the credential file
    $credential = Get-Credential -Message "Enter your username and password"
    $credential | Export-Clixml -Path $credentialFile
}

# Import the CSV file and connect to each appliance
Import-Csv -Path $csvFilePath | ForEach-Object {
    Connect-OneViewAppliance -ApplianceFQDN $_.FQDN -Credential $credential
}