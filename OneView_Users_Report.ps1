$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$loggingFunctionsPath = Join-Path -Path $scriptPath -ChildPath "..\Logging_Function\Logging_Functions.ps1"
. $loggingFunctionsPath
$ScriptVersion = "1.1"
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
# Function to connect to an appliance
Function Connect-OneViewAppliance {
    param (
        [string]$ApplianceIP,
        [string]$Username,
        [SecureString]$Password
    )
    $OneViewCreds = New-Object System.Management.Automation.PSCredential($Username, ($Password | ConvertTo-SecureString -AsPlainText -Force))
    Connect-OVMgmt -Hostname $ApplianceIP -Credential $OneViewCreds
}

# Import the CSV file
$AppliancesList = Import-Csv -Path $csvFilePath

# Loop through each appliance and connect
foreach ($appliance in $AppliancesList) {
    Connect-OneViewAppliance -ApplianceIP $appliance.ApplianceIP -Username $appliance.Username -Password $appliance.Password
}


