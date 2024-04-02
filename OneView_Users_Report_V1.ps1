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
# Second Task import the CSV file
Write-Host "`n$Spaces$($taskNumber). Importing the CSV file:`n" -ForegroundColor Cyan
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
# Initialize arrays to hold local users and LDAP groups
$allLocalUsers = @()
$allLdapGroups = @()
# Define the path to the Excel file for local users
$localUsersExcelPath = Join-Path -Path $script:ReportsDir -ChildPath 'LocalUsers.xlsx'
# Define the path to the Excel file for LDAP groups
$ldapGroupsExcelPath = Join-Path -Path $script:ReportsDir -ChildPath 'LdapGroups.xlsx'
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
# Define the Convert-ToLetter function
function Convert-ToLetter {
    param (
        [int]$column
    )

    $alphabet = , @(65..90)
    $columnName = while ($column) {
        $digit = $column % 26
        $column = [math]::Floor($column / 26)
        $alphabet[$digit]
    }
    -join [char[]]$columnName
}
# Export the local users to an Excel file
$allLocalUsers | Export-Excel -Path $localUsersExcelPath
# Export the LDAP groups to an Excel file
$allLdapGroups | Export-Excel -Path $ldapGroupsExcelPath
# Select specific properties from local users and add LDAP group-specific properties with default values
$selectedLocalUsers = $allLocalUsers | Select-Object ApplianceConnection, type, category, userName, fullName, Role, @{Name = 'loginDomain'; Expression = { 'NO' } }, @{Name = 'egroup'; Expression = { 'N/A' } }, @{Name = 'directoryType'; Expression = { 'User' } }
# Select specific properties from LDAP groups and add local user-specific properties with default values
$selectedLdapGroups = $allLdapGroups | Select-Object ApplianceConnection, type, category, @{Name = 'userName'; Expression = { 'N/A' } }, @{Name = 'fullName'; Expression = { 'N/A' } }, Role, loginDomain, egroup, directoryType
# Combine selected local users and LDAP groups into a single array
$selectedUsers = $selectedLocalUsers + $selectedLdapGroups
# Define the path to the Excel file for combined user details
$combinedUsersExcelPath = Join-Path -Path $script:ReportsDir -ChildPath 'CombinedUsers.xlsx'
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
    Write-Host "has been closed.`n" -ForegroundColor Green
    Write-Log "The file '$(Split-Path $filePath -Leaf)' has been closed." -Level "OK" -NoConsoleOutput
}
# Call the function
Close-ExcelFile -filePath $combinedUsersExcelPath
# Sort the selected user details based on ApplianceConnection
$sortedUsers = $selectedUsers | Sort-Object -Property ApplianceConnection
# Export the sorted user details to an Excel file
$sortedUsers | Export-Excel -Path $combinedUsersExcelPath -AutoSize -FreezeTopRow
# Open the Excel package
$excel = Open-ExcelPackage -Path $combinedUsersExcelPath
# Check if a worksheet named 'Users_details' already exists
if ($excel.Workbook.Worksheets.Name -contains 'Users_details') {
    # If it exists, delete it
    $excel.Workbook.Worksheets.Delete('Users_details')
}
# Rename the first worksheet to 'Users_details'
$worksheet = $excel.Workbook.Worksheets[1]
$worksheet.Name = 'Users_details'

# Get the number of properties in the selected users
$propertyCount = ($selectedUsers | Get-Member -MemberType NoteProperty).Count

if ($propertyCount -eq 0) {
    Write-Host "No properties found in selected users"
    return
}

# Convert the property count to a column letter
$propertyCountLetter = Convert-ToLetter $propertyCount

if (-not $propertyCountLetter) {
    Write-Host "Failed to convert property count to letter"
    return
}

# Apply formatting to the headers of the selected properties
$range = $worksheet.Cells["A1:$($propertyCountLetter)1"]
$range.Style.Font.Bold = $true
$range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$range.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::DarkBlue)
$range.Style.Font.Color.SetColor([System.Drawing.Color]::White)

# Save and close the Excel package
Close-ExcelPackage $excel -Show
# Just before calling Complete-Logging
$endTime = Get-Date
$totalRuntime = $endTime - $startTime
# Call Complete-Logging at the end of the script
Complete-Logging -LogPath $script:LogPath -ErrorCount $ErrorCount -WarningCount $WarningCount -TotalRuntime $totalRuntime
