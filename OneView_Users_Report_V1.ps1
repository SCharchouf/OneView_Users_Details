# Clear the console window
Clear-Host
# Define the script version
$ScriptVersion = "1.0"
# Get the directory from which the script is being executed
$scriptDirectory = $PSScriptRoot
# Get the parent directory of the script's directory
$parentPath = Split-Path -Parent $scriptDirectory
# Define the logging function Directory
$loggingFunctionsDirectory = Join-Path -Path $parentPath -ChildPath "Logging_Function"
# Construct the path to the Logging_Functions.ps1 script
$loggingFunctionsPath = Join-Path -Path $loggingFunctionsDirectory -ChildPath "Logging_Functions.ps1"
# Script Header main script
$HeaderMainScript = @"
Author : Your Name
Description : This script does amazing things!
Created : $(Get-Date -Format "dd/MM/yyyy")
Last Modified : $((Get-Item $PSCommandPath).LastWriteTime.ToString("dd/MM/yyyy"))
"@
# Display the header information in the console with a design
$consoleWidth = $Host.UI.RawUI.WindowSize.Width
$line = "─" * ($consoleWidth - 2)
Write-Host "+$line+" -ForegroundColor Cyan
# Split the header into lines and display each part in different colors
$HeaderMainScript -split "`n" | ForEach-Object {
    $parts = $_ -split ": ", 2
    Write-Host "`t" -NoNewline
    Write-Host $parts[0] -NoNewline -ForegroundColor DarkGray
    Write-Host ": " -NoNewline
    Write-Host $parts[1] -ForegroundColor Cyan
}
Write-Host "+$line+" -ForegroundColor Cyan
# Check if the Logging_Functions.ps1 script exists
if (Test-Path -Path $loggingFunctionsPath) {
    # Dot-source the Logging_Functions.ps1 script
    . $loggingFunctionsPath
    # Write a message to the console indicating that the logging functions have been loaded
    Write-Host "`t• Logging functions have been loaded." -ForegroundColor Green
}
else {
    # Write an error message to the console indicating that the logging functions script could not be found
    Write-Host "`t• The logging functions script could not be found at: $loggingFunctionsPath" -ForegroundColor Red
    # Stop the script execution
    exit
}
# Initialize task counter
$script:taskNumber = 1
# Define the function to import required modules if they are not already imported
function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ModuleNames
    )
    # Start logging
    Start-Log -ScriptVersion $ScriptVersion -ScriptPath $PSCommandPath
    # Task 1: Checking required modules
    Write-Host "`n$($taskNumber). Checking required modules:`n" -ForegroundColor Magenta
    # Log the task
    Write-Log -Message "Checking required modules." -Level "Info" -NoConsoleOutput
    # Increment $script:taskNumber after the function call
    $script:taskNumber++
    # Total number of modules to check
    $totalModules = $ModuleNames.Count
    # Initialize the current module counter
    $currentModuleNumber = 0
    # Loop through each module name
    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        Write-Progress -Activity "Checking required modules" -Status "$ModuleName" -PercentComplete ($currentModuleNumber / $totalModules * 100)
        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t• Module " -NoNewline -ForegroundColor White
                Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Red
                Write-Host " is not installed." -ForegroundColor White
                Write-Log -Message "Module '[$ModuleName]' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }
            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t• Module " -NoNewline -ForegroundColor White
                Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Yellow
                Write-Host " is already imported." -ForegroundColor White
                Write-Log -Message "Module '[$ModuleName]' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }
            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t• Module " -NoNewline -ForegroundColor White
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Green
            Write-Host " imported successfully." -ForegroundColor White
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t• Failed to import module " -NoNewline
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
# Define the CSV file name
$csvFileName = ".\Appliances_List\Appliances_List.csv"
# Define the parent directory of the CSV file
$parentDirectory = Split-Path -Path $scriptDirectory -Parent
# Create the full path to the CSV file
$csvFilePath = Join-Path -Path $parentDirectory -ChildPath $csvFileName
# Define the path to the credential folder
$credentialFolder = Join-Path -Path $scriptDirectory -ChildPath "credential"
# Define the path to the credential file
$credentialFile = Join-Path -Path $credentialFolder -ChildPath "credential.txt"
# Second Task import the CSV file
Write-Host "`n$($taskNumber). Importing the CSV file:`n" -ForegroundColor Magenta
# Import Appliances list from CSV file
$Appliances = Import-Csv -Path $csvFilePath
# Confirm that the CSV file was imported successfully
if ($Appliances) {
    # Get the total number of appliances
    $totalAppliances = $Appliances.Count
    # Log the total number of appliances
    Write-Log -Message "There are $totalAppliances appliances in the CSV file." -Level "Info" -NoConsoleOutput
    # Display if the CSV file was imported successfully
    Write-Host "`t• The CSV file was imported successfully." -ForegroundColor Green
    # Display the total number of appliances
    Write-Host "`t• Total number of appliances:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $totalAppliances" -NoNewline -ForegroundColor Cyan
    Write-Host "" # This is to add a newline after the above output
    # Log the successful import of the CSV file
    Write-Log -Message "The CSV file was imported successfully." -Level "OK" -NoConsoleOutput
}
else {
    # Display an error message if the CSV file failed to import
    Write-Host "`t• Failed to import the CSV file." -ForegroundColor Red
    # Log the failure to import the CSV file
    Write-Log -Message "Failed to import the CSV file." -Level "Error" -NoConsoleOutput
}
# Increment $script:taskNumber after the function call
$script:taskNumber++
# Third Task : Check for existing sessions
Write-Host "`n$($taskNumber). Check for existing sessions:`n" -ForegroundColor Magenta 
$existingSessions = $ConnectedSessions
if ($existingSessions) {
    Write-Host "`t• Existing sessions found: $($existingSessions.Count)" -ForegroundColor Yellow
    Write-Log -Message "Existing sessions found: $($existingSessions.Count)" -Level "Info" -NoConsoleOutput
    # Disconnect all existing sessions
    $existingSessions | ForEach-Object {
        Disconnect-OVMgmt -Hostname $_
    }
    Write-Host "`t• All existing sessions have been disconnected." -ForegroundColor Green
    Write-Log -Message "All existing sessions have been disconnected." -Level "OK" -NoConsoleOutput
}
else {
    Write-Host "`t• No existing sessions found." -ForegroundColor Green
    Write-Log -Message "No existing sessions found." -Level "Info" -NoConsoleOutput
}
# Increment $script:taskNumber after the function call
$script:taskNumber++
# Fourth Task : Loop through each appliance
Write-Host "`n$($taskNumber). Loop through each appliance & Collect users details:`n" -ForegroundColor Magenta 
# Create an empty array to store the user details
$userDetails = @()
# Check if the credential file exists
if (-not (Test-Path -Path $credentialFile)) {
    # Prompt the user to enter their login and password
    $credential = Get-Credential -Message "Please enter your login and password."
    # Save the credential to the credential file
    $credential | Export-Clixml -Path $credentialFile
}
else {
    # Load the credential from the credential file
    $credential = Import-Clixml -Path $credentialFile
}
# Loop through each appliance
foreach ($appliance in $Appliances) {
    # Convert the FQDN to uppercase
    $fqdn = $appliance.FQDN.ToUpper()

    try {
        # Use the Connect-OVMgmt cmdlet to connect to the appliance
        Connect-OVMgmt -Hostname $fqdn -Credential $credential | Out-Null

        Write-Host "`t• Successfully connected to :" -NoNewline -ForegroundColor Green
        Write-Host " $fqdn" -ForegroundColor Blue
        Write-Log -Message "Successfully connected to $fqdn." -Level "OK" -NoConsoleOutput

        # Here, you can add the code to collect user details or perform other tasks

        # Disconnect from the appliance
        Disconnect-OVMgmt -Hostname $fqdn
    }
    catch {
        # Log the failed connection
        Write-Log -Message "Failed to connect to appliance: $fqdn. Error: $($_.Exception.Message)" -Level "Error" -NoConsoleOutput
    }
}
# Define the path to the Excel file in the Reports directory
$excelFilePath = Join-Path -Path $script:reportsDir -ChildPath 'UserDetails.xlsx'
# Export the user details to an Excel file
$userDetails | Export-Excel -Path $excelFilePath
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime
