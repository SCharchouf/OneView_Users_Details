# Define the script version
$ScriptVersion = "1.0"
# Get the directory from which the script is being executed
$scriptDirectory = $PSScriptRoot
# Move up one level from the current script directory and then into the Logging_Function directory
$loggingFunctionsDirectory = Join-Path -Path $scriptDirectory -ChildPath "..\Logging_Function"
# Construct the path to the Logging_Functions.ps1 script
$loggingFunctionsPath = Join-Path -Path $loggingFunctionsDirectory -ChildPath "Logging_Functions.ps1"
# Script Header
# Get the width of the console window
$consoleWidth = $host.UI.RawUI.WindowSize.Width
# Create a string of "=" characters that is as long as the console is wide
$line = "=" * ($consoleWidth - 1)
# Script Header
Write-Host "`n$line" -ForegroundColor Cyan
Write-Host "  MY SCRIPT" -ForegroundColor Yellow
Write-Host "  Version: $ScriptVersion" -ForegroundColor Yellow
Write-Host "  Description: This script does amazing things!" -ForegroundColor Yellow
Write-Host "$line`n" -ForegroundColor Cyan
# Check if the Logging_Functions.ps1 script exists
if (Test-Path -Path $loggingFunctionsPath) {
    # Dot-source the Logging_Functions.ps1 script
    . $loggingFunctionsPath
    Write-Host "Logging functions have been loaded."
}
else {
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
    Write-Host "`n$($taskNumber). Checking required modules:`n" -ForegroundColor Magenta
    Write-Log -Message "Checking required modules." -Level "Info"
    $taskNumber++

    $totalModules = $ModuleNames.Count
    $currentModuleNumber = 0

    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        Write-Progress -Activity "Checking required modules" -Status "$ModuleName" -PercentComplete ($currentModuleNumber / $totalModules * 100)

        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t* Module " -NoNewline -ForegroundColor White
                Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Red
                Write-Host " is not installed." -ForegroundColor White
                Write-Log -Message "Module '[$ModuleName]' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }

            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t* Module " -NoNewline -ForegroundColor White
                Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Yellow
                Write-Host " is already imported." -ForegroundColor White
                Write-Log -Message "Module '[$ModuleName]' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }

            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t* Module " -NoNewline -ForegroundColor White
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Green
            Write-Host " imported successfully." -ForegroundColor White
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t* Failed to import module " -NoNewline
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Red
            Write-Host ": $_" -ForegroundColor Red
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

