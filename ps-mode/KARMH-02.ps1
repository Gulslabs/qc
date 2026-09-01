$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
$logPath = Join-Path $scriptRoot 'error.logs'

function Write-ErrorLog([string]$Message) {
    Add-Content -Path $logPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

try {
    Set-Location $scriptRoot
    $paramsPath = Join-Path $scriptRoot 'params.info'
    if (-not (Test-Path $paramsPath)) { throw 'params.info was not found.' }
    $config = @{}
    Get-Content $paramsPath | ForEach-Object {
        if ($_ -match '^\s*([^#=]+?)\s*=\s*"?(.*?)"?\s*$') { $config[$matches[1].Trim()] = $matches[2].Trim() }
    }
    foreach ($key in @('start_time', 'end_time', 'bins', 'delete_non_naqeeb_records')) {
        if (-not $config.ContainsKey($key)) { throw "Missing '$key' in params.info." }
    }

    $csvFiles = @(Get-ChildItem -Path $scriptRoot -Filter 'participant*.csv' -File)
    if ($csvFiles.Count -eq 0) { throw 'No participant*.csv file was found in ps-mode. Add exactly one participant CSV file.' }
    if ($csvFiles.Count -gt 1) { throw "More than one participant*.csv file was found. Keep only one file: $($csvFiles.Name -join ', ')" }

    $python = if (Get-Command python.exe -ErrorAction SilentlyContinue) { 'python' } elseif (Get-Command py.exe -ErrorAction SilentlyContinue) { 'py' } else { throw 'Python is not installed. Run setup.ps1 first.' }
    $converter = Join-Path $scriptRoot '..\02_zoom_participant_csv_to_timesheet_extractor\csv_to_timesheet_extractor_v3.py'
    $appender = Join-Path $scriptRoot '..\03_timesheet_report_appender\timesheet_report_appender_v1.py'
    $mapping = Join-Path $scriptRoot '..\02_zoom_participant_csv_to_timesheet_extractor\naqeeb_to_initial_KARMH-02.csv'
    $before = @(Get-ChildItem -Path $scriptRoot -Filter '*.xlsx' -File | Select-Object Name, LastWriteTime)
    $started = Get-Date

    & $python $converter $csvFiles[0].FullName $mapping $config['start_time'] $config['end_time']
    if ($LASTEXITCODE -ne 0) { throw 'csv_to_timesheet_extractor_v3.py failed.' }
    $timesheet = Get-ChildItem -Path $scriptRoot -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '*_with_report.xlsx' -and $_.LastWriteTime -ge $started } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $timesheet) { throw 'The CSV converter did not create an Excel output file.' }

    & $python $appender $timesheet.FullName ([int]$config['bins'])
    if ($LASTEXITCODE -ne 0) { throw 'timesheet_report_appender_v1.py failed.' }
    $report = Join-Path $scriptRoot ($timesheet.BaseName + '_with_report.xlsx')
    if (-not (Test-Path $report)) { throw 'The report appender did not create its Excel output file.' }
    Write-Host "Completed successfully. Outputs are in ps-mode:`n  $($timesheet.Name)`n  $([System.IO.Path]::GetFileName($report))"
} catch {
    Write-ErrorLog $_.Exception.Message
    Write-ErrorLog $_.ScriptStackTrace
    Write-ErrorLog 'KARMH-02 process terminated.'
    Write-Error "Process failed. Details were recorded in $logPath"
    exit 1
}
