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
    # Initialize task counter
    $taskNumber = 1

    # Task 1: Checking required modules
    Write-Host "`n$($taskNumber). Checking required modules:`n" -ForegroundColor Cyan
    $taskNumber++

    $totalModules = $ModuleNames.Count
    $currentModuleNumber = 0

    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        Write-Progress -Activity "Checking required modules" -Status "$ModuleName" -PercentComplete ($currentModuleNumber / $totalModules * 100)

        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t- Module '[$ModuleName]' is not installed." -ForegroundColor Red
                Write-Log -Message "Module '[$ModuleName]' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }

            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t- Module '[$ModuleName]' is already imported." -ForegroundColor Yellow
                Write-Log -Message "Module '[$ModuleName]' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }

            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t- Module '[$ModuleName]' imported successfully." -ForegroundColor Green
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t- Failed to import module '[$ModuleName]': $_" -ForegroundColor Red
            Write-Log -Message "Failed to import module '[$ModuleName]': $_" -Level "Error" -NoConsoleOutput
        }

        # Add a delay to slow down the progress bar
        Start-Sleep -Seconds 1
    }
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

