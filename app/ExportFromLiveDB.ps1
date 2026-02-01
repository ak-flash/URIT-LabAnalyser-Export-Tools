<#
.SYNOPSIS
    Exports patient and result data from the LIVE URIT Chemistry Analyzer database to Excel (XLSX) and CSV.
    Portable version - works from USB or any folder.
    Includes logging to 'export_log.txt'.
#>

# Ensure correct encoding for Console Output
$OutputEncoding = [Console]::InputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- Load Configuration ---
$ScriptPath = $PSScriptRoot
$RootPath = Split-Path -Parent $ScriptPath
$ConfigFile = $RootPath + '\config.json'

if (-not (Test-Path $ConfigFile)) {
    Write-Host ('Error: Configuration file not found at ' + $ConfigFile) -ForegroundColor Red
    exit 1
}

try {
    $config = Get-Content -Path $ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
    $dbHost = $config.dbHost
    $dbName = $config.dbName
    $dbUser = $config.dbUser
    $dbPwd  = $config.dbPwd
    
    # Handle enableLogging with robust type checking
    if ($null -ne $config.enableLogging) {
        # Convert to int first to handle "0" string or 0 number correctly, then to bool
        # In PowerShell [bool]"0" is True, but [bool]0 is False.
        if ($config.enableLogging -is [string] -and $config.enableLogging -eq '0') {
            $enableLogging = $false
        }
        else {
            $enableLogging = [bool]$config.enableLogging
        }
    } else {
        $enableLogging = $true
    }
} catch {
    Write-Host ('Error reading configuration file: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
}

$ExportFolder = $RootPath + '\ExportedResults'
$LogFile = $RootPath + '\export_log.txt'

# Ensure Export Folder Exists
if (-not (Test-Path $ExportFolder)) {
    New-Item -ItemType Directory -Path $ExportFolder | Out-Null
}

# --- Logging Function ---
function Log-Message {
    param(
        [string]$Message,
        [string]$Color = 'White'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = '[' + $timestamp + '] ' + $Message
    
    # Write to Console
    Write-Host $Message -ForegroundColor $Color
    
    # Write to File
    if ($enableLogging) {
        try {
            Add-Content -Path $LogFile -Value $logEntry -Encoding UTF8
        } catch {
            # Ignore logging errors
        }
    }
}

# Clear previous log if it is a new run session and logging is enabled
if ($enableLogging -and (Test-Path $LogFile)) {
    try {
        Add-Content -Path $LogFile -Value '----------------------------------------' -Encoding UTF8
    } catch {}
}

Log-Message -Message '=== URIT Live Data Exporter (Portable) Started ===' -Color 'Cyan'
if (-not $enableLogging) {
    Log-Message -Message 'Logging to file is DISABLED in config.json' -Color 'Yellow'
}
Log-Message -Message ('Script Path: ' + $ScriptPath)
Log-Message -Message ('Target Server: ' + $dbHost)
Log-Message -Message ('Target Database: ' + $dbName)
Log-Message -Message ('Target User: ' + $dbUser)

# --- Date Selection Logic ---
$todayStr = Get-Date -Format 'dd-MM-yyyy'

Write-Host ('Введите дату для экспорта (формат dd-MM-yyyy) [По умолчанию: ' + $todayStr + ']: ') -ForegroundColor Cyan -NoNewline
$defaultDateStr = Read-Host

if ([string]::IsNullOrWhiteSpace($defaultDateStr)) {
    $targetDateStr = $todayStr
} else {
    $targetDateStr = $defaultDateStr
}

# Convert dd-MM-yyyy to yyyyMMdd for SQL filtering (ID starts with yyyyMMdd)
try {
    $parsedDate = [DateTime]::ParseExact($targetDateStr, 'dd-MM-yyyy', $null)
    $sqlDateFilter = $parsedDate.ToString('yyyyMMdd')
} catch {
    Log-Message -Message ('Неверный формат даты! Используется текущая дата: ' + $todayStr) -Color 'Red'
    $sqlDateFilter = (Get-Date).ToString('yyyyMMdd')
    $targetDateStr = $todayStr
}

# --- Define SQL Function ---
function Run-SqlQuery {
    param(
        [string]$Query,
        [bool]$ReturnData = $false
    )
    
    # Build Connection String
    $connString = 'Server=' + $dbHost + ';Database=' + $dbName + ';User Id=' + $dbUser + ';Password=' + $dbPwd + ';'
    
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connString
    
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = 300
        
        if ($ReturnData) {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
            $dataset = New-Object System.Data.DataSet
            [void]$adapter.Fill($dataset)
            return ,$dataset.Tables[0]
        } else {
            [void]$command.ExecuteNonQuery()
        }
    }
    catch {
        throw $_
    }
    finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
    }
}

# --- Helper Function: Export Query to CSV ---
function Export-QueryToCsv {
    param(
        [string]$Query,
        [string]$FileName,
        [string]$Description
    )

    Log-Message -Message ('Exporting ' + $Description + '...') -Color 'White'
    try {
        $dt = Run-SqlQuery -Query $Query -ReturnData $true
        
        # Handle potential array wrapping
        if ($dt -is [Array]) {
            $dtFound = $dt | Where-Object { $_ -is [System.Data.DataTable] } | Select-Object -First 1
            if ($dtFound) { 
                $dt = $dtFound 
            }
            elseif ($dt[0] -is [System.Data.DataRow]) {
                 $dt = $dt[0].Table
            }
        }

        if ($dt -eq $null) { 
             Log-Message -Message '  -> Warning: Query returned no data (DataTable is null).' -Color 'Yellow'
             return
        }
        
        Log-Message -Message ('  Rows to process: ' + $dt.Rows.Count) -Color 'Gray'
        
        $csvPath = $ExportFolder + '\' + $FileName
        
        if ($dt.Rows.Count -gt 0) {
            $columns = $dt.Columns
            
            $objList = $dt.Rows | ForEach-Object {
                $row = $_
                $obj = New-Object PSObject
                
                foreach ($col in $columns) {
                    $val = $row.Item($col.ColumnName)
                    if ($val -is [DBNull]) { 
                        $val = '' 
                    }
                    $obj | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value $val
                }
                $obj
            }
            
            # Export to CSV with UTF-8 NoBOM to ensure compatibility with external tools
            $csvData = $objList | ConvertTo-Csv -NoTypeInformation
            [System.IO.File]::WriteAllLines($csvPath, $csvData, (New-Object System.Text.UTF8Encoding $false))
            
            Log-Message -Message ('  -> Success: Saved to ' + $csvPath) -Color 'Green'
        }
        else {
            Log-Message -Message '  -> Warning: Result is empty, skipping export.' -Color 'Yellow'
        }
    }
    catch {
        Log-Message -Message ('  -> Error exporting ' + $Description + ' : ' + $_.Exception.Message) -Color 'Red'
    }
}

# --- Helper Function: Convert CSV to Excel (XLSX) using csv2xlsx.exe ---
function Convert-CsvToExcel {
    param (
        [string]$CsvPath
    )
    
    $ExcelPath = $CsvPath -replace '\.csv$', '.xlsx'
    # csv2xlsx.exe is expected to be in the same folder as the script
    $ToolPath = $ScriptPath + '\csv2xlsx.exe'
    
    if (-not (Test-Path $ToolPath)) {
        Log-Message -Message ('  -> Error: csv2xlsx.exe not found at ' + $ToolPath) -Color 'Red'
        return
    }

    Log-Message -Message ('Converting to Excel using csv2xlsx: ' + $ExcelPath + '...') -Color 'White'
    
    if (Test-Path $ExcelPath) {
        Remove-Item $ExcelPath -Force
    }

    try {
        # csv2xlsx usage: csv2xlsx -o output.xlsx input.csv
        # Construct arguments carefully to avoid parser issues
        $q = '"'
        $argString = '-o ' + $q + $ExcelPath + $q + ' ' + $q + $CsvPath + $q
        
        $process = Start-Process -FilePath $ToolPath -ArgumentList $argString -NoNewWindow -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Log-Message -Message ('  -> Success: Converted to Excel: ' + $ExcelPath) -Color 'Green'
        } else {
            Log-Message -Message ('  -> Error converting to Excel. Exit code: ' + $process.ExitCode) -Color 'Red'
        }
    }
    catch {
        Log-Message -Message ('  -> Error executing csv2xlsx: ' + $_.Exception.Message) -Color 'Red'
    }
}

# --- Helper Function: Format Excel File (Adjust Widths) ---
function Format-ExcelFile {
    param (
        [string]$ExcelPath
    )
    
    Log-Message -Message 'Applying formatting to Excel file...' -Color 'White'
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $absExcelPath = (Resolve-Path $ExcelPath).Path
        $workbook = $excel.Workbooks.Open($absExcelPath)
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Adjust Column Widths
        # Column 1: CheckDate - Set to 15
        $worksheet.Columns.Item(1).ColumnWidth = 10
        # Column 3: PatientName - Set to 30 (Wide)
        $worksheet.Columns.Item(3).ColumnWidth = 30
        
        $workbook.Save()
        $workbook.Close()
        
        if ($worksheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
        if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null }
        
        # Log-Message -Message '  -> Success: Formatted Excel columns (CheckDate and PatientName widths increased).' -Color 'Green'
    }
    catch {
        Log-Message -Message '  -> Warning: Could not apply Excel formatting. Ensure Microsoft Excel is installed.' -Color 'Yellow'
        Log-Message -Message ('  -> Details: ' + $_.Exception.Message) -Color 'Gray'
    }
    finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
    }
}

# --- Test Connection ---
Log-Message -Message 'Подключение к базе данных...' -Color 'Yellow'
try {
    Run-SqlQuery -Query 'SELECT @@VERSION' | Out-Null
    Log-Message -Message 'Подключение успешно!' -Color 'Green'
}
catch {
    Log-Message -Message 'FATAL ERROR: Failed to connect to database.' -Color 'Red'
    Log-Message -Message ('Error details: ' + $_.Exception.Message) -Color 'Red'
    Log-Message -Message 'Ensure the URIT software/SQL Server is running on this computer.' -Color 'Yellow'
    Log-Message -Message '=== Export Failed ===' -Color 'Red'
    exit 1
}

Log-Message -Message ('Проверка наличия данных за дату: ' + $targetDateStr + ' (' + $sqlDateFilter + ')...') -Color 'White'

# Check if data exists
$checkQuery = 'SELECT COUNT(*) FROM PATIENT_DATABASE WHERE ID LIKE ''' + $sqlDateFilter + '%'''
try {
    $dtCount = Run-SqlQuery -Query $checkQuery -ReturnData $true
    
    if ($dtCount -is [System.Data.DataTable] -and $dtCount.Rows.Count -gt 0) {
        $count = $dtCount.Rows[0][0]
    }
    elseif ($dtCount -is [Array] -and $dtCount.Count -gt 0) {
         $first = $dtCount[0]
         if ($first -is [System.Data.DataRow]) { $count = $first[0] }
         else { $count = 0 }
    }
    else {
        $count = 0
    }
    
    if ($count -eq 0) {
        Log-Message -Message ('Данные за ' + $targetDateStr + ' не найдены.') -Color 'Yellow'
        
        $maxDateQuery = 'SELECT TOP 1 LEFT(ID, 8) as MaxDate FROM PATIENT_DATABASE WHERE ISNUMERIC(LEFT(ID, 8)) = 1 ORDER BY ID DESC'
        $dtMax = Run-SqlQuery -Query $maxDateQuery -ReturnData $true
        
        if ($dtMax.Rows.Count -gt 0) {
            $latestDateRaw = $dtMax.Rows[0]["MaxDate"]
            if (-not [string]::IsNullOrWhiteSpace($latestDateRaw)) {
                $sqlDateFilter = $latestDateRaw
                try {
                    $latestDateParsed = [DateTime]::ParseExact($latestDateRaw, 'yyyyMMdd', $null)
                    $targetDateStr = $latestDateParsed.ToString('dd-MM-yyyy')
                } catch {
                    $targetDateStr = $latestDateRaw
                }
                Log-Message -Message ('Переключение на последнюю доступную дату: ' + $targetDateStr) -Color 'Red'
            } else {
                 Log-Message -Message 'В базе данных вообще нет подходящих записей.' -Color 'Red'
            }
        }
    } else {
        Log-Message -Message ('Найдены записи: ' + $count) -Color 'Green'
    }
} catch {
    Log-Message -Message ('Ошибка при проверке данных: ' + $_.Exception.Message) -Color 'Red'
}

# --- Export Consolidated Report (JOIN) ---
$filter = $sqlDateFilter + '%'

# Use explicit string concatenation and single quotes to avoid parser confusion
$reportQuery = 'SELECT ' +
    'SUBSTRING(p.ID, 7, 2) + ''-'' + SUBSTRING(p.ID, 5, 2) + ''-'' + LEFT(p.ID, 4) AS CheckDate, ' +
    'SUBSTRING(p.ID, 9, LEN(p.ID)) AS ResultNumber, ' +
    'p.FIRST_NAME AS PatientName, ' +
    'CASE WHEN TRY_CAST(p.AGE AS FLOAT) = 0 THEN NULL ELSE p.AGE END AS BirthYear, ' +
    'c.ITEM AS TestName, ' +
    'c.RESULT AS TestResult, ' +
    'c.UNIT AS TestUnit, ' +
    'p.DOCTOR AS Doctor ' +
    'FROM PATIENT_DATABASE p ' +
    'JOIN check_result c ON p.ID = c.ID ' +
    'WHERE p.ID LIKE ''' + $filter + ''' ' +
    'ORDER BY p.ID, c.ITEM'

Export-QueryToCsv -Query $reportQuery `
                  -FileName ($targetDateStr + '_Full_Report_Patients_Results.csv') `
                  -Description 'Consolidated Report (Patients + Results)'

# Convert to Excel
$csvFullPath = $ExportFolder + '\' + $targetDateStr + '_Full_Report_Patients_Results.csv'
if (Test-Path $csvFullPath) {
    Convert-CsvToExcel -CsvPath $csvFullPath
    
    # Apply post-processing formatting if possible
    $xlsxFullPath = $csvFullPath -replace '\.csv$', '.xlsx'
    if (Test-Path $xlsxFullPath) {
        Format-ExcelFile -ExcelPath $xlsxFullPath
    }
}

Log-Message -Message ('=== Finished! Files are in: ' + $ExportFolder + ' ===') -Color 'Green'
if ($enableLogging) {
    Log-Message -Message 'Check export_log.txt for full log.' -Color 'Cyan'
}

[Environment]::Exit(0)
