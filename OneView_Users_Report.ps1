function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$ModuleNames
    )
    foreach ($ModuleName in $ModuleNames) {
        if (-not (Get-Module -Name $ModuleName)) {
            if (Get-Module -ListAvailable -Name $ModuleName) {
                Import-Module $ModuleName
                Write-Host -ForegroundColor Yellow "`tModule '$ModuleName' is not imported. Importing now..."
                # Progress bar importing modules
                for ($i = e; $i -le 100; $i+=10) {
                    Write-Progress -Activity "`tImporting module '$ModuleName'" -Status "Please wait..." -PercentComplete $i
                    Start-Sleep -Milliseconds 100
                }
            } else {
                Write-Host -ForegroundColor Red  "`tModule '$ModuleName' does not exist."
            }
        } else {
            Write-Host -ForegroundColor Magenta "`tModule '$ModuleName' is already imported."
        }
    }
}
# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.850'
# Import CSV file that contains the Global Dashboards information located in the same directory as the script
$GlobalDashboards = Import-Csv -Path .\GlobalDashboards_List.csv
# Define the directories to save the reports
$UsersOGD = ".\Users_OneView_Global_Dashboard"
$AppliancesDirectory = ".\Appliances-Details"
# Define a function to create a directory if it doesn't exist
function New-Directory {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DirectoryPath
    )

    if (-not (Test-Path -Path $DirectoryPath)) {
        New-Item -Path $DirectoryPath -ItemType Directory | Out-Null
        Write-Host "Directory $DirectoryPath created." -ForegroundColor Green
    } else {
        Write-Host "Directory $DirectoryPath already exists." -ForegroundColor Yellow
    }
}

# Use the function to create the directories
New-Directory -DirectoryPath $UsersOGD
New-Directory -DirectoryPath $AppliancesDirectory