# Clear the console window
Clear-Host
# Create a string of 4 spaces
$Spaces = [string]::new(' ', 4)
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
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Logging functions have been loaded." -ForegroundColor Green
}
else {
    # Write an error message to the console indicating that the logging functions script could not be found
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The logging functions script could not be found at: $loggingFunctionsPath" -ForegroundColor Red
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
    Write-Host "`n$Spaces$($taskNumber). Checking required modules:`n" -ForegroundColor Cyan
    # Log the task
    Write-Log -Message "Checking required modules." -Level "Info" -NoConsoleOutput
    # Increment $script:taskNumber after the function call
    $script:taskNumber++
    # Total number of modules to check
    $totalModules = $ModuleNames.Count
    # Initialize the current module counter
    $currentModuleNumber = 0
    foreach ($ModuleName in $ModuleNames) {
        $currentModuleNumber++
        # Simple text output for checking required modules
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Checking module " -NoNewline -ForegroundColor DarkGray
        Write-Host "$currentModuleNumber" -NoNewline -ForegroundColor White
        Write-Host " of " -NoNewline -ForegroundColor DarkGray
        Write-Host "${totalModules}" -NoNewline -ForegroundColor Cyan
        Write-Host ": $ModuleName" -ForegroundColor White
        try {
            # Check if the module is installed
            if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor White
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Red
                Write-Host " is not installed." -ForegroundColor White
                Write-Log -Message "Module '$ModuleName' is not installed." -Level "Error" -NoConsoleOutput
                continue
            }
            # Check if the module is already imported
            if (Get-Module -Name $ModuleName) {
                Write-Host "`t• " -NoNewline -ForegroundColor White
                Write-Host "Module " -NoNewline -ForegroundColor DarkGray
                Write-Host "$ModuleName" -NoNewline -ForegroundColor Yellow
                Write-Host " is already imported." -ForegroundColor DarkGray
                Write-Log -Message "Module '$ModuleName' is already imported." -Level "Info" -NoConsoleOutput
                continue
            }
            # Try to import the module
            Import-Module $ModuleName -ErrorAction Stop
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Module " -NoNewline -ForegroundColor DarkGray
            Write-Host "[$ModuleName]" -NoNewline -ForegroundColor Green
            Write-Host " imported successfully." -ForegroundColor DarkGray
            Write-Log -Message "Module '[$ModuleName]' imported successfully." -Level "OK" -NoConsoleOutput
        }
        catch {
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "Failed to import module " -NoNewline
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
# Second Task import Appliances list from the CSV file and loop through each appliance to collect user details.
Write-Host "`n$Spaces$($taskNumber). Importing Appliances list from the CSV file:`n" -ForegroundColor Cyan
# Import Appliances list from CSV file
$Appliances = Import-Csv -Path $csvFilePath
# Confirm that the CSV file was imported successfully
if ($Appliances) {
    # Get the total number of appliances
    $totalAppliances = $Appliances.Count
    # Log the total number of appliances
    Write-Log -Message "There are $totalAppliances appliances in the CSV file." -Level "Info" -NoConsoleOutput
    # Display if the CSV file was imported successfully
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The CSV file was imported successfully." -ForegroundColor Green
    # Display the total number of appliances
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Total number of appliances:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $totalAppliances" -NoNewline -ForegroundColor Cyan
    Write-Host "" # This is to add a newline after the above output
    # Log the successful import of the CSV file
    Write-Log -Message "The CSV file was imported successfully." -Level "OK" -NoConsoleOutput
}
else {
    # Display an error message if the CSV file failed to import
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Failed to import the CSV file." -ForegroundColor Red
    # Log the failure to import the CSV file
    Write-Log -Message "Failed to import the CSV file." -Level "Error" -NoConsoleOutput
}
# increment $script:taskNumber after the function call
$script:taskNumber++
# Third Task : Loop through each appliance
Write-Host "`n$Spaces$($taskNumber). Loop through each appliance & Collect users details:`n" -ForegroundColor Cyan
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
# Initialize arrays
$allLocalUsers = @()
$allLdapGroups = @()
# Define the directories for the CSV and Excel files
$csvDir = Join-Path -Path $script:ReportsDir -ChildPath 'CSV'
$excelDir = Join-Path -Path $script:ReportsDir -ChildPath 'Excel'
# Check if the CSV directory exists
if (Test-Path -Path $csvDir) {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $csvDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "CSV directory already exists at $csvDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory does not exist." -NoNewline -ForegroundColor Red
    Write-Host " Creating now..." -ForegroundColor DarkGray
    Write-Log -Message "CSV directory does not exist, creating now..." -Level "Info" -NoConsoleOutput
    # Create the CSV directory if it does not exist already
    New-Item -ItemType Directory -Path $csvDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "CSV directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $csvDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "CSV directory created at $csvDir" -Level "OK" -NoConsoleOutput
}
# Check if the Excel directory exists
if (Test-Path -Path $excelDir) {
    # Write a message to the console
    write-host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory already exists at:" -NoNewline -ForegroundColor DarkGray
    write-host " $excelDir" -ForegroundColor Yellow
    # Write a message to the log file
    Write-Log -Message "Excel directory already exists at $excelDir" -Level "Info" -NoConsoleOutput
}
else {
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory does not exist at" -NoNewline -ForegroundColor Red
    Write-Host " $excelDir" -ForegroundColor DarkGray
    # Write a message to the log file
    Write-Log -Message "Excel directory does not exist at $excelDir, creating now..." -Level "Info" -NoConsoleOutput
    # Create the Excel directory if it does not exist already
    New-Item -ItemType Directory -Path $excelDir | Out-Null
    # Write a message to the console
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "Excel directory created at:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $excelDir" -ForegroundColor Green
    # Write a message to the log file
    Write-Log -Message "Excel directory created at $excelDir" -Level "OK" -NoConsoleOutput
}
# Define the path to the CSV and Excel files for local users and LDAP groups
$localUsersCsvPath = Join-Path -Path $csvDir -ChildPath 'LocalUsers.csv'
$ldapGroupsCsvPath = Join-Path -Path $csvDir -ChildPath 'LdapGroups.csv'
$localUsersExcelPath = Join-Path -Path $excelDir -ChildPath 'LocalUsers.xlsx'
$ldapGroupsExcelPath = Join-Path -Path $excelDir -ChildPath 'LdapGroups.xlsx'
$combinedUsersExcelPath = Join-Path -Path $excelDir -ChildPath 'CombinedUsers.xlsx'
# Loop through each appliance
foreach ($appliance in $Appliances) {
    # Convert the FQDN to uppercase
    $fqdn = $appliance.FQDN.ToUpper()
    # Check for existing sessions and disconnect them
    $existingSessions = $ConnectedSessions
    if ($existingSessions) {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "Existing sessions found: $($existingSessions.Count)" -ForegroundColor Yellow
        Write-Log -Message "Existing sessions found: $($existingSessions.Count)" -Level "Info" -NoConsoleOutput
        # Disconnect all existing sessions
        $existingSessions | ForEach-Object {
            Disconnect-OVMgmt -Hostname $_
        }
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "All existing sessions have been disconnected." -ForegroundColor Green
        Write-Log -Message "All existing sessions have been disconnected." -Level "OK" -NoConsoleOutput
        # Add a small delay to ensure the session is fully disconnected
        Start-Sleep -Seconds 5
    }
    else {
        Write-Host "`t• " -NoNewline -ForegroundColor White
        Write-Host "No existing sessions found.`n" -ForegroundColor Gray
        Write-Log -Message "No existing sessions found." -Level "Info" -NoConsoleOutput
    }
    # Use the Connect-OVMgmt cmdlet to connect to the appliance
    Connect-OVMgmt -Hostname $fqdn -Credential $credential *> $null
    Write-Host "`t1- Successfully connected to:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    Write-Log -Message "Successfully connected to: $fqdn" -Level "OK" -NoConsoleOutput
    # Collect user details
    Write-Host "`t2- Collecting user details from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn"  -ForegroundColor Green
    $users = Get-OVUser | ForEach-Object {
        $_ | Add-Member -NotePropertyName 'Role' -NotePropertyValue ($_.permissions | ForEach-Object { $_.roleName }) -PassThru
    }
    $allLocalUsers += $users
    # Collect LDAP group details
    Write-Host "`t3- Collecting LDAP group details from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    $ldapGroups = Get-OVLdapGroup | ForEach-Object {
        $_ | Add-Member -NotePropertyName 'Role' -NotePropertyValue ($_.permissions | ForEach-Object { $_.roleName }) -PassThru
    }
    $allLdapGroups += $ldapGroups
    # Generate reports
    Write-Host "`t4- Generating report for:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn" -ForegroundColor Green
    # Disconnect from the appliance
    Disconnect-OVMgmt -Hostname $fqdn
    Write-Host "`t5- Successfully disconnected from:" -NoNewline -ForegroundColor DarkGray
    Write-Host " $fqdn`n" -ForegroundColor Green
    Write-Log -Message "Successfully disconnected from $fqdn" -Level "OK" -NoConsoleOutput
}
# increment $script:taskNumber after the function call
$script:taskNumber++
# Fourth Task : Assembling the Excel file
Write-Host "`n$Spaces$($taskNumber). Assembling the Excel file:`n" -ForegroundColor Cyan
# Export the local users to an Excel file
$allLocalUsers | Export-Excel -Path $localUsersExcelPath
# Export the LDAP groups to an Excel file
$allLdapGroups | Export-Excel -Path $ldapGroupsExcelPath
# Assign the local users and LDAP groups to variables
$allLocalUsersCsv = $allLocalUsers
$allLdapGroupsCsv = $allLdapGroups
# Export the local users to a CSV file
$allLocalUsersCsv | Export-Csv -Path $localUsersCsvPath -NoTypeInformation
# Export the LDAP groups to a CSV file
$allLdapGroupsCsv | Export-Csv -Path $ldapGroupsCsvPath -NoTypeInformation
# Select specific properties from local users and add LDAP group-specific properties with default values
$selectedLocalUsers = $allLocalUsersCsv | Select-Object ApplianceConnection, type, category, userName, fullName, Role, @{Name = 'loginDomain'; Expression = { 'NO' } }, @{Name = 'egroup'; Expression = { 'N/A' } }, @{Name = 'directoryType'; Expression = { 'User' } }
# Select specific properties from LDAP groups and add local user-specific properties with default values
$selectedLdapGroups = $allLdapGroupsCsv | Select-Object ApplianceConnection, type, category, @{Name = 'userName'; Expression = { 'N/A' } }, @{Name = 'fullName'; Expression = { 'N/A' } }, Role, loginDomain, egroup, directoryType
# Combine all local users and LDAP groups into a single array
$combinedUsers = $selectedLocalUsers + $selectedLdapGroups
# Define Close-ExcelFile function to close the Excel file if it is open 
function Close-ExcelFile {
    param (
        [string]$filePath
    )
    # Check if the file is open
    $delay = 10
    while ((Test-Path $filePath) -and (Get-Process excel -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*$(Split-Path $filePath -Leaf)*" })) {
        try {
            # Write a message to the console
            Write-Host "`t• " -NoNewline -ForegroundColor White
            Write-Host "The file " -NoNewline -ForegroundColor Yellow
            Write-Host "'$(Split-Path $filePath -Leaf)'" -NoNewline -ForegroundColor Cyan
            Write-Host " is currently open. Attempting to close it..." -ForegroundColor Yellow
            # Write the message to a log file
            Write-Log -Message "The file '$(Split-Path $filePath -Leaf)' is currently open. Attempting to close it..." -Level 'Warning' -NoConsoleOutput
            # Attempt to close the Excel file
            $excelProcess = Get-Process excel | Where-Object { $_.MainWindowTitle -like "*$(Split-Path $filePath -Leaf)*" }
            $excelProcess | ForEach-Object { $_.CloseMainWindow() | Out-Null }
            # Wait for a moment to ensure the process has time to close
            Start-Sleep -Seconds $delay
            $delay = [math]::max(1, $delay - 1)
        }
        catch {
            Write-Error "An error occurred while trying to close the Excel file: $_"
        }
    }
    Write-Host "`t• " -NoNewline -ForegroundColor White
    Write-Host "The file " -NoNewline -ForegroundColor DarkGray
    Write-Host "'$(Split-Path $filePath -Leaf)'" -NoNewline -ForegroundColor Cyan
    Write-Host " has been closed.`n" -ForegroundColor Green
    Write-Log "The file '$(Split-Path $filePath -Leaf)' has been closed." -Level "OK" -NoConsoleOutput
}
# Call the function
Close-ExcelFile -filePath $combinedUsersExcelPath
# Sort the combined users by ApplianceConnection and then by userName
$sortedCombinedUsers = $combinedUsers | Sort-Object ApplianceConnection, userName
# Export the sorted user details to an Excel file
$sortedCombinedUsers | Export-Excel -Path $combinedUsersExcelPath -AutoSize -FreezeTopRow -AutoFilter -WorkSheetname "CombinedUsers" -TabColor Yellow -PassThru
# ------------------------------------------------------------
function Convert-ToColumnName($number) {
    $columnName = ""
    while ($number -gt 0) {
        $mod = ($number - 1) % 26
        $columnName = [char](65 + $mod) + $columnName
        $number = [math]::Floor(($number - $mod) / 26)
    }
    return $columnName
}
# Get the number of properties
$propertyCount = ($sortedCombinedUsers | Get-Member -MemberType NoteProperty).Count
write-host "Property Count: $propertyCount"

# Convert the number of properties to the corresponding Excel column letter
$columnLetter = Convert-ToColumnName $propertyCount

# Construct the range for the title row
$titleRowRange = "A1:$columnLetter" + "1"

# Format the title row
$excel = $sortedCombinedUsers | Export-Excel -Path $combinedUsersExcelPath -AutoSize -FreezeTopRow -AutoFilter -WorkSheetname "CombinedUsers" -TabColor Yellow -PassThru
$ws = $excel.Workbook.Worksheets["CombinedUsers"]
$ws.Cells[$titleRowRange].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$ws.Cells[$titleRowRange].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Blue)
$ws.Cells[$titleRowRange].Style.Font.Color.SetColor([System.Drawing.Color]::White)
$ws.Cells[$titleRowRange].Style.Font.Size = 12
$ws.Cells[$titleRowRange].Style.Font.Bold = $true
$excel.Save()
$excel.Dispose()
# ------------------------------------------------------------
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime
