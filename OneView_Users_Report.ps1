function Import-ModulesIfNotExists {
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$ModuleNames
    )
    $totalModules = $ModuleNames.Count
    $currentModule = 0
    foreach ($ModuleName in $ModuleNames) {
        if (-not (Get-Module -Name $ModuleName)) {
            if (Get-Module -ListAvailable -Name $ModuleName) {
                Import-Module $ModuleName
                Write-Host "`tModule '$ModuleName' is not imported. Importing now..." -ForegroundColor Yellow
                $currentModule++
                $percentComplete = ($currentModule / $totalModules) * 100
                Write-Progress -Activity "`tImporting modules" -Status "Imported module '$ModuleName'" -PercentComplete $percentComplete
            } else {
                Write-Host "`tModule '$ModuleName' does not exist." -ForegroundColor Red
            }
        } else {
            Write-Host "`tModule '$ModuleName' is already imported." -ForegroundColor Green
        }
    }
}
# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.850', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility'