# Define the script version
$ScriptVersion = "1.0"
# Retrieve the script name dynamically
$ScriptName = $MyInvocation.MyCommand.Name
# Get the directory from which the script is being executed
$scriptDirectory = $PSScriptRoot

# Move up one level from the current script directory and then into the Logging_Function directory
$loggingFunctionsDirectory = Join-Path -Path $scriptDirectory -ChildPath "..\Logging_Function"

# Construct the path to the Logging_Functions.ps1 script
$loggingFunctionsPath = Join-Path -Path $loggingFunctionsDirectory -ChildPath "Logging_Functions.ps1"

# Check if the Logging_Functions.ps1 script exists
if (Test-Path -Path $loggingFunctionsPath) {
    # Dot-source the Logging_Functions.ps1 script
    . $loggingFunctionsPath
    Write-Host "Logging functions have been loaded."
} else {
    Write-Host "The logging functions script could not be found at: $loggingFunctionsPath"
}
# Define the function to import required modules if they are not already imported
function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ModuleNames
    )
    # Start logging
    Start-Log -ScriptVersion $ScriptVersion -ScriptPath $PSCommandPath
    # Add space before the progress bar
    Write-Host "`nChecking required modules:`n"

    $totalModules = $ModuleNames.Count
    $currentModuleNumber = 0

    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        Write-Progress -Activity "Checking required modules" -Status "$ModuleName" -PercentComplete ($currentModuleNumber / $totalModules * 100)

        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Log -Message "Module '[$ModuleName]' is not installed." -Level "Error"
                continue
            }

            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Log -Message "Module '[$ModuleName]' is already imported." -Level "Info"
                continue
            }

            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK"
        }
        catch {
            Write-Log -Message "Failed to import module '[$ModuleName]': $_" -Level "Error"
        }
    }

    Write-Host "`n"
}

# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.660', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility', 'ImportExcel'
# Create the full path to the CSV file
$csvFilePath = Join-Path -Path $scriptDirectory -ChildPath $csvFileName
# Define the path to the credential folder
$credentialFolder = Join-Path -Path $scriptDirectory -ChildPath "credential"
# Define the path to the credential file
$credentialFile = Join-Path -Path $credentialFolder -ChildPath "credential.txt"
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime

