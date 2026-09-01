$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
$logPath = Join-Path $scriptRoot 'error.logs'

function Write-ErrorLog([string]$Message) {
    Add-Content -Path $logPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

try {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $elevatedProcess = Start-Process powershell.exe -Verb RunAs -PassThru -Wait -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
        exit $elevatedProcess.ExitCode
    }

    Set-Location $scriptRoot
    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    $launcherCommand = Get-Command py.exe -ErrorAction SilentlyContinue
    $python = $null

    if ($pythonCommand) {
        $python = 'python'
    } elseif ($launcherCommand) {
        $python = 'py'
    } else {
        Write-Host 'Python was not found. Downloading the official Windows installer...'
        $installerPath = Join-Path $env:TEMP 'python-3.13.7-amd64.exe'
        wget -Uri 'https://www.python.org/ftp/python/3.13.7/python-3.13.7-amd64.exe' -OutFile $installerPath
        Start-Process $installerPath -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_launcher=1' -Wait
        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
        $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
        $launcherCommand = Get-Command py.exe -ErrorAction SilentlyContinue
        if ($pythonCommand) { $python = 'python' } elseif ($launcherCommand) { $python = 'py' }
    }

    if (-not $python) { throw 'Python installation was not found after setup.' }
    & $python --version
    if ($LASTEXITCODE -ne 0) { throw 'Python was found but could not be started.' }

    $missingPackages = @()
    foreach ($package in @('pandas', 'numpy', 'openpyxl')) {
        & $python -c "import $package" 2>$null
        if ($LASTEXITCODE -ne 0) { $missingPackages += $package }
    }

    if ($missingPackages.Count -gt 0) {
        Write-Host "Installing missing packages: $($missingPackages -join ', ')"
        & $python -m pip install -r (Join-Path $scriptRoot '..\requirements.txt')
        if ($LASTEXITCODE -ne 0) { throw 'Package installation failed.' }
    } else {
        Write-Host 'All required Python packages are already installed.'
    }

    Write-Host 'Setup completed successfully. You can now run KARMH-02.ps1, HYDLH-01.ps1, or TSAP-02.ps1.'
} catch {
    Write-ErrorLog $_.Exception.Message
    Write-ErrorLog $_.ScriptStackTrace
    Write-ErrorLog 'Setup terminated.'
    Write-Error "Setup failed. Details were recorded in $logPath"
    exit 1
}
