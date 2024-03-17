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
# Define folder path where Global Dashboards are stored. Floder called Global_Dashboards_List, on the same level as the folder from where the script is run.
$GlobalDashboardsPath = ".\Global_Dashboards_List"
# Import GDashboards_List in CSV from the folder "Global_Dashboards_List"
$GlobalDashboardsList = Import-Csv -Path "$GlobalDashboardsPath\GDashboards_List.csv"
# Create a new folder to store the reports
$ReportPath = ".\Reports"
if (-not (Test-Path -Path $ReportPath)) {
    New-Item -Path $ReportPath -ItemType Directory
}
 
