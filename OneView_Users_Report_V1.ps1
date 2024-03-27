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
                    Write-Log -Message $message -Level "Error"
                }
                else {
                    $message = "`tModule '$ModuleName' imported successfully."
                    Write-Log -Message $message -Level "OK"
                }
            }
            else {
                $message = "`tModule '$ModuleName' is already imported."
                Write-Log -Message $message -Level "Info"
            }
        }
        else {
            $message = "`tModule '$ModuleName' does not exist."
            Write-Log -Message $message -Level "Error" -Path $script:LogPath
        }
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
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime -FinalStatus $FinalStatus

