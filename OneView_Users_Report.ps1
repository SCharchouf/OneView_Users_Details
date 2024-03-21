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

    $totalModules = $ModuleNames.Count
    $currentModule = 0
    foreach ($ModuleName in $ModuleNames) {
        if (-not (Get-Module -Name $ModuleName)) {
            if (Get-Module -ListAvailable -Name $ModuleName) {
                Import-Module $ModuleName
                $message = "`tModule '$ModuleName' is not imported. Importing now..."
                Write-Log -Message $message -Level "Warning" -sFullPath $global:sFullPath
                $currentModule++
                $percentComplete = ($currentModule / $totalModules) * 100
                Write-Progress -Activity "`tImporting modules" -Status "Imported module '$ModuleName'" -PercentComplete $percentComplete
                $message = "`tModule '$ModuleName' imported successfully."
                Write-Log -Message $message -Level "OK" -sFullPath $global:sFullPath
            } else {
                $message = "`tModule '$ModuleName' does not exist."
                Write-Log -Message $message -Level "Error" -sFullPath $global:sFullPath
            }
        } else {
            $message = "`tModule '$ModuleName' is already imported."
            Write-Log -Message $message -Level "Info" -sFullPath $global:sFullPath
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

Function Connect-OneViewAppliance {
    param (
        [string]$ApplianceFQDN,
        [PSCredential]$Credential
    )

    try {
        # Check if a connection to the appliance already exists
        $existingConnection = $ConnectedSessions | Where-Object { $_.Hostname -eq $ApplianceFQDN }

        if ($existingConnection) {
            # If a connection already exists, write and log a message
            $message = "Already connected to : $ApplianceFQDN"
            Write-Host $message
            Write-Log -Message $message -Level "Info" -sFullPath $global:sFullPath
        }

        # Always attempt to connect to the appliance
        $connection = Connect-OVMgmt -Hostname $ApplianceFQDN -Credential $Credential

        # If the connection is successful, write and log a success message
        if ($connection) {
            $message = "Successfully connected to : $ApplianceFQDN"
            Write-Host $message
            Write-Log -Message $message -Level "OK" -sFullPath $global:sFullPath
        }
    } catch {
        Write-Log -Message "Failed to connect to : $ApplianceFQDN. Error details: $_" -Level "Error" -sFullPath $global:sFullPath
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