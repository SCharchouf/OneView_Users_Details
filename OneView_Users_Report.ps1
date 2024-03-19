function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level,
        [Parameter(Mandatory=$false)]
        [string]$Path,
        [Parameter(Mandatory=$false)]
        [string]$ModuleName
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Level - $Message"
    Add-Content -Path $logFilePath -Value $logMessage

    switch ($Level) {
        "Error" { 
            if ($ModuleName) {
                Write-Host "`tModule '" -NoNewline
                Write-Host $ModuleName -NoNewline -ForegroundColor Cyan
                Write-Host "' does not exist." -ForegroundColor Red
            } else {
                Write-Host $Message -ForegroundColor Red
            }
        }
        "Warn"  { 
            if ($ModuleName) {
                Write-Host "`tModule '" -NoNewline
                Write-Host $ModuleName -NoNewline -ForegroundColor Cyan
                Write-Host "' is not imported. Importing now..." -ForegroundColor Yellow
            } else {
                Write-Host $Message -ForegroundColor Yellow
            }
        }
        "Info"  { 
            if ($ModuleName) {
                Write-Host "`tModule '" -NoNewline
                Write-Host $ModuleName -NoNewline -ForegroundColor Cyan
                Write-Host "' is already imported." -ForegroundColor Green
            } else {
                Write-Host $Message -ForegroundColor Green
            }
        }
    }
}
$scriptName = (Get-Item $MyInvocation.MyCommand.Definition).BaseName
$logDirPath = Join-Path $PSScriptRoot ($scriptName + "_LOG")

if (!(Test-Path $logDirPath)) {
    New-Item -ItemType Directory -Path $logDirPath -Force
    Write-Log -Message "Log directory $logDirPath created." -Level Info
} else {
    Write-Log -Message "Log directory $logDirPath already exists." -Level Info
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logFileName = $scriptName + "_" + $timestamp + ".log"
$logFilePath = Join-Path $logDirPath $logFileName

if (!(Test-Path $logFilePath)) {
    New-Item -ItemType File -Path $logFilePath -Force
    Write-Log -Message "Log file $logFilePath created." -Level Info
} else {
    Write-Log -Message "Log file $logFilePath already exists." -Level Info
}
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
                $message = "`tModule '$ModuleName' is not imported. Importing now..."
                Write-Log -Message $message -Level "Warn" -Path $logFilePath
                $currentModule++
                $percentComplete = ($currentModule / $totalModules) * 100
                Write-Progress -Activity "`tImporting modules" -Status "Imported module '$ModuleName'" -PercentComplete $percentComplete
            } else {
                $message = "`tModule '$ModuleName' does not exist."
                Write-Log -Message $message -Level "Error" -Path $logFilePath
            }
        } else {
            $message = "`tModule '$ModuleName' is already imported."
            Write-Log -Message $message -Level "Info" -Path $logFilePath
        }
    }
}
# Import the required modules
Import-ModulesIfNotExists -ModuleNames 'HPEOneView.850', 'Microsoft.PowerShell.Security', 'Microsoft.PowerShell.Utility'