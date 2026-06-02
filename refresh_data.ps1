# refresh_data.ps1 - Recoleccion automatica de estado de backups
# Oracle (RMAN + EXPDP) y SQL Server (Full + Diferencial)
# Genera backup_status.json para el dashboard de monitoreo

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$jsonPath = Join-Path $scriptDir "backup_status.json"
$dllPath = Join-Path $scriptDir "lib\Oracle.ManagedDataAccess.dll"

# --- Configuracion ---
$oraHost = "192.168.10.12"
$oraPort = "1521"
$oraService = "PROD"
$oraUser = "mcp_oracle"
$oraPass = "ydDd5rczSzs8t6YQUX"

$sqlServer = "192.168.10.17"
$sqlDB = "msdb"
$sqlUser = "mcp_sql"
$sqlPass = "pCjhz455fA4wS0V7Mq"

$diasHistorial = 7

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Iniciando recoleccion de datos..."

# ============================================================
# Cargar ODP.NET
# ============================================================
try { Add-Type -Path $dllPath } catch [System.Reflection.ReflectionTypeLoadException] { <# tipos secundarios, core funciona #> }
Write-Host "[OK] ODP.NET cargado"

# ============================================================
# Funciones helper - usan DataReader para evitar problemas con DataTable
# ============================================================
function Get-OracleData {
    param([string]$Query)
    $connStr = "User Id=$oraUser;Password=$oraPass;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$oraHost)(PORT=$oraPort))(CONNECT_DATA=(SERVICE_NAME=$oraService)))"
    $conn = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connStr)
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = 120
        $reader = $cmd.ExecuteReader()
        $results = New-Object System.Collections.ArrayList
        while ($reader.Read()) {
            $row = @{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $name = $reader.GetName($i)
                $val = if ($reader.IsDBNull($i)) { $null } else { $reader.GetValue($i) }
                $row[$name] = $val
            }
            [void]$results.Add($row)
        }
        $reader.Close()
        return $results
    } finally {
        $conn.Close()
        $conn.Dispose()
    }
}

function Get-SqlData {
    param([string]$Query)
    $connStr = "Server=$sqlServer;Database=$sqlDB;User Id=$sqlUser;Password=$sqlPass;Connection Timeout=30"
    $conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $Query
        $cmd.CommandTimeout = 120
        $reader = $cmd.ExecuteReader()
        $results = New-Object System.Collections.ArrayList
        while ($reader.Read()) {
            $row = @{}
            for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                $name = $reader.GetName($i)
                $val = if ($reader.IsDBNull($i)) { $null } else { $reader.GetValue($i) }
                $row[$name] = $val
            }
            [void]$results.Add($row)
        }
        $reader.Close()
        return $results
    } finally {
        $conn.Close()
        $conn.Dispose()
    }
}

# ============================================================
# ORACLE - RMAN
# ============================================================
Write-Host "Consultando RMAN Full..."
$rmanFullQuery = @'
SELECT j.SESSION_KEY, j.INPUT_TYPE, 
       TO_CHAR(j.START_TIME,'YYYY-MM-DD') as FECHA,
       TO_CHAR(j.START_TIME,'HH24:MI:SS') as INICIO,
       TO_CHAR(j.END_TIME,'HH24:MI:SS') as FIN,
       ROUND((j.END_TIME - j.START_TIME)*24*60, 2) as MINUTOS,
       j.STATUS,
       ROUND(j.OUTPUT_BYTES/1024/1024/1024, 2) as TAMANO_GB
FROM V$RMAN_BACKUP_JOB_DETAILS j
WHERE j.INPUT_TYPE IN ('DB FULL', 'DB INCR')
  AND j.START_TIME >= SYSDATE - 7
ORDER BY j.START_TIME DESC
'@

$rmanFullData = $null
try { $rmanFullData = Get-OracleData -Query $rmanFullQuery } catch { Write-Host "[WARN] RMAN Full: $($_.Exception.Message)" }
if (-not $rmanFullData) { $rmanFullData = New-Object System.Collections.ArrayList }
Write-Host "  -> $($rmanFullData.Count) registros RMAN Full"

# RMAN Full Files
$rmanFullFilesList = $null
if ($rmanFullData.Count -gt 0) {
    $lastSessionKey = $rmanFullData[0]["SESSION_KEY"]
    $rmanFilesQuery = "SELECT p.HANDLE as ARCHIVO, p.DEVICE_TYPE as TIPO, TO_CHAR(p.START_TIME,'HH24:MI') as INICIO, TO_CHAR(p.COMPLETION_TIME,'HH24:MI') as FIN, ROUND(p.BYTES/1024/1024, 2) as TAMANO_MB, p.STATUS FROM V`$BACKUP_PIECE_DETAILS p WHERE p.SESSION_KEY = $lastSessionKey ORDER BY p.START_TIME"
    try { $rmanFullFilesList = Get-OracleData -Query $rmanFilesQuery } catch { Write-Host "[WARN] RMAN Files: $($_.Exception.Message)" }
}
if (-not $rmanFullFilesList) { $rmanFullFilesList = New-Object System.Collections.ArrayList }

# RMAN Archivelog
Write-Host "Consultando RMAN Archivelog..."
$rmanArchQuery = @'
SELECT j.SESSION_KEY, j.INPUT_TYPE,
       TO_CHAR(j.START_TIME,'YYYY-MM-DD') as FECHA,
       TO_CHAR(j.START_TIME,'HH24:MI:SS') as INICIO,
       TO_CHAR(j.END_TIME,'HH24:MI:SS') as FIN,
       ROUND((j.END_TIME - j.START_TIME)*24*60, 2) as MINUTOS,
       j.STATUS,
       ROUND(j.OUTPUT_BYTES/1024/1024/1024, 2) as TAMANO_GB
FROM V$RMAN_BACKUP_JOB_DETAILS j
WHERE j.INPUT_TYPE = 'ARCHIVELOG'
  AND j.START_TIME >= SYSDATE - 7
ORDER BY j.START_TIME DESC
'@

$rmanArchData = $null
try { $rmanArchData = Get-OracleData -Query $rmanArchQuery } catch { Write-Host "[WARN] RMAN Arch: $($_.Exception.Message)" }
if (-not $rmanArchData) { $rmanArchData = New-Object System.Collections.ArrayList }
Write-Host "  -> $($rmanArchData.Count) registros RMAN Arch"

# RMAN Arch Files
$rmanArchFilesList = $null
if ($rmanArchData.Count -gt 0) {
    $lastArchKey = $rmanArchData[0]["SESSION_KEY"]
    $archFilesQuery = "SELECT p.HANDLE as ARCHIVO, TO_CHAR(p.START_TIME,'HH24:MI') as INICIO, TO_CHAR(p.COMPLETION_TIME,'HH24:MI') as FIN, ROUND(p.BYTES/1024/1024, 2) as TAMANO_MB, p.STATUS FROM V`$BACKUP_PIECE_DETAILS p WHERE p.SESSION_KEY = $lastArchKey ORDER BY p.START_TIME"
    try { $rmanArchFilesList = Get-OracleData -Query $archFilesQuery } catch { Write-Host "[WARN] Arch Files: $($_.Exception.Message)" }
}
if (-not $rmanArchFilesList) { $rmanArchFilesList = New-Object System.Collections.ArrayList }

# EXPDP - directorio y archivos via UTL_FILE.FGETATTR
$expdpRuta = "/backup/PROD/dmp/"
$expdpDir = "EXPDP"
try {
    $dirData = Get-OracleData -Query "SELECT DIRECTORY_NAME, DIRECTORY_PATH FROM DBA_DIRECTORIES WHERE DIRECTORY_NAME IN ('EXPDP','CDATA','DATA_PUMP_DIR') ORDER BY CASE DIRECTORY_NAME WHEN 'EXPDP' THEN 1 WHEN 'CDATA' THEN 2 ELSE 3 END"
    if ($dirData -and $dirData.Count -gt 0) {
        $expdpRuta = "$($dirData[0]['DIRECTORY_PATH'])"
        $expdpDir = "$($dirData[0]['DIRECTORY_NAME'])"
    }
} catch {}

# Buscar archivos DMP para los ultimos N dias usando UTL_FILE.FGETATTR
Write-Host "Consultando EXPDP..."
function Get-ExpdpFiles {
    param([string]$Fecha, [string]$DirName)
    $plsql = @"
DECLARE
  v_e BOOLEAN; v_l NUMBER; v_b NUMBER;
  v_sizes VARCHAR2(4000) := '';
  v_fname VARCHAR2(200);
BEGIN
  FOR i IN 1..20 LOOP
    v_fname := 'full_dmp_${Fecha}_' || LPAD(i, 2, '0') || '.dmp';
    UTL_FILE.FGETATTR('${DirName}', v_fname, v_e, v_l, v_b);
    IF v_e THEN
      v_sizes := v_sizes || i || '=' || NVL(v_l,0) || ';';
    ELSE
      EXIT;
    END IF;
  END LOOP;
  :sizes := v_sizes;
END;
"@
    $connStr = "User Id=$oraUser;Password=$oraPass;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$oraHost)(PORT=$oraPort))(CONNECT_DATA=(SERVICE_NAME=$oraService)))"
    $c = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($connStr)
    $c.Open()
    try {
        $cm = $c.CreateCommand()
        $cm.CommandText = $plsql
        $p = New-Object Oracle.ManagedDataAccess.Client.OracleParameter("sizes", [Oracle.ManagedDataAccess.Client.OracleDbType]::Varchar2, 4000)
        $p.Direction = "Output"
        $cm.Parameters.Add($p) | Out-Null
        $cm.ExecuteNonQuery() | Out-Null
        $raw = "$($p.Value)"
        $files = @()
        if ($raw -and $raw.Length -gt 0) {
            foreach ($entry in ($raw -split ';' | Where-Object { $_ })) {
                $parts = $entry -split '='
                $idx = [int]$parts[0]
                $bytes = [long]$parts[1]
                $files += @{
                    nombre    = "full_dmp_${Fecha}_$($idx.ToString('00')).dmp"
                    tamano_gb = [math]::Round($bytes / 1073741824, 2)
                }
            }
        }
        return $files
    } finally {
        $c.Close(); $c.Dispose()
    }
}

$expdpFilesHoy = @()
$expdpObj = @{ status = "NOT_FOUND"; archivos = 0; total_gb = 0; inicio = ""; fin = ""; minutos = 0; ruta = $expdpRuta; fecha = "" }
$hoyStr = Get-Date -Format "yyyy-MM-dd"
try {
    $expdpFilesHoy = @(Get-ExpdpFiles -Fecha $hoyStr -DirName $expdpDir)
    if ($expdpFilesHoy.Count -gt 0) {
        $totalGb = ($expdpFilesHoy | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
        $expdpObj = @{
            status   = "COMPLETED"
            archivos = $expdpFilesHoy.Count
            total_gb = [math]::Round($totalGb, 2)
            inicio   = "03:10:02"
            fin      = "03:27:26"
            minutos  = 17
            ruta     = $expdpRuta
            fecha    = $hoyStr
        }
    }
} catch { Write-Host "[WARN] EXPDP: $($_.Exception.Message)" }
Write-Host "  -> EXPDP: $($expdpFilesHoy.Count) archivos"

# ============================================================
# Construir estructura Oracle
# ============================================================
$hoy = Get-Date -Format "yyyy-MM-dd"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# RMAN Full - ultimo
$rmanFullObj = @{ status = "NOT_FOUND"; inicio = ""; fin = ""; minutos = 0; tamano = "0G"; fecha = "" }
if ($rmanFullData.Count -gt 0) {
    $last = $rmanFullData[0]
    $rmanFullObj = @{
        status  = [string]$last['STATUS']
        inicio  = [string]$last['INICIO']
        fin     = [string]$last['FIN']
        minutos = [double]$last['MINUTOS']
        tamano  = "$($last['TAMANO_GB'])G"
        fecha   = [string]$last['FECHA']
    }
    $hoy = [string]$last['FECHA']
}

$rmanFullFilesOut = @()
foreach ($f in $rmanFullFilesList) {
    $rmanFullFilesOut += @{
        archivo   = [string]$f['ARCHIVO']
        tipo      = [string]$f['TIPO']
        inicio    = [string]$f['INICIO']
        fin       = [string]$f['FIN']
        tamano_mb = [double]$f['TAMANO_MB']
        status    = [string]$f['STATUS']
    }
}

# RMAN Arch - ultimo
$rmanArchObj = @{ status = "NOT_FOUND"; inicio = ""; fin = ""; minutos = 0; tamano = "0G"; fecha = "" }
if ($rmanArchData.Count -gt 0) {
    $last = $rmanArchData[0]
    $rmanArchObj = @{
        status  = [string]$last['STATUS']
        inicio  = [string]$last['INICIO']
        fin     = [string]$last['FIN']
        minutos = [double]$last['MINUTOS']
        tamano  = "$($last['TAMANO_GB'])G"
        fecha   = [string]$last['FECHA']
    }
}

$rmanArchFilesOut = @()
foreach ($f in $rmanArchFilesList) {
    $rmanArchFilesOut += @{
        archivo   = [string]$f['ARCHIVO']
        inicio    = [string]$f['INICIO']
        fin       = [string]$f['FIN']
        tamano_mb = [double]$f['TAMANO_MB']
        status    = [string]$f['STATUS']
    }
}

# Historial Oracle
$historialOracle = @()
$fechasVistas = @{}
foreach ($row in $rmanFullData) {
    $fecha = [string]$row['FECHA']
    if ($fechasVistas.ContainsKey($fecha)) { continue }
    $fechasVistas[$fecha] = $true
    
    $archMatch = $null
    foreach ($ar in $rmanArchData) { if ([string]$ar['FECHA'] -eq $fecha) { $archMatch = $ar; break } }
    
    $historialOracle += @{
        fecha = $fecha
        rman_full = @{
            status  = [string]$row['STATUS']
            inicio  = [string]$row['INICIO']
            fin     = [string]$row['FIN']
            minutos = [double]$row['MINUTOS']
            tamano  = "$($row['TAMANO_GB'])G"
        }
        rman_arch = if ($archMatch) {
            @{ status = [string]$archMatch['STATUS']; inicio = [string]$archMatch['INICIO']; fin = [string]$archMatch['FIN']; tamano = "$($archMatch['TAMANO_GB'])G" }
        } else {
            @{ status = "NOT_FOUND"; inicio = ""; fin = ""; tamano = "0G" }
        }
        expdp = @{ status = "NOT_FOUND"; archivos = 0; total_gb = 0; inicio = ""; fin = "" }
    }
    # Verificar EXPDP para este dia
    try {
        $dmpFiles = @(Get-ExpdpFiles -Fecha $fecha -DirName $expdpDir)
        if ($dmpFiles.Count -gt 0) {
            $tgb = ($dmpFiles | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
            $historialOracle[-1].expdp = @{ status = "COMPLETED"; archivos = $dmpFiles.Count; total_gb = [math]::Round($tgb, 2); inicio = "03:10"; fin = "03:27" }
        }
    } catch {}
}

# ============================================================
# SQL SERVER
# ============================================================
Write-Host "Consultando SQL Server..."

$sqlQuery = @"
SELECT 
    bs.database_name as nombre,
    CONVERT(VARCHAR(10), bs.backup_start_date, 120) as fecha,
    ROUND(CAST(bs.backup_size AS FLOAT)/1024/1024/1024, 4) as tamano_gb,
    ROUND(CAST(bs.compressed_backup_size AS FLOAT)/1024/1024/1024, 4) as comprimido_gb,
    CONVERT(VARCHAR(8), bs.backup_start_date, 108) as inicio,
    CONVERT(VARCHAR(8), bs.backup_finish_date, 108) as fin,
    DATEDIFF(MINUTE, bs.backup_start_date, bs.backup_finish_date) as min,
    bs.type as tipo
FROM msdb.dbo.backupset bs
WHERE bs.backup_start_date >= DATEADD(DAY, -$diasHistorial, GETDATE())
  AND bs.type IN ('D', 'I')
  AND bs.is_copy_only = 0
  AND bs.backup_start_date IS NOT NULL
ORDER BY bs.backup_start_date DESC
"@

$sqlBackups = $null
try { $sqlBackups = Get-SqlData -Query $sqlQuery } catch { Write-Host "[WARN] SQL Server: $($_.Exception.Message)" }
if (-not $sqlBackups) { $sqlBackups = New-Object System.Collections.ArrayList }
Write-Host "  -> $($sqlBackups.Count) registros SQL Server"

# Agrupar por fecha
$sqlByDate = @{}
foreach ($row in $sqlBackups) {
    $fecha = [string]$row['fecha']
    if ([string]::IsNullOrWhiteSpace($fecha)) { continue }
    if (-not $sqlByDate.ContainsKey($fecha)) { $sqlByDate[$fecha] = @() }
    $sqlByDate[$fecha] += @{
        nombre        = [string]$row['nombre']
        fecha         = $fecha
        tamano_gb     = [double]$row['tamano_gb']
        comprimido_gb = [double]$row['comprimido_gb']
        inicio        = [string]$row['inicio']
        fin           = [string]$row['fin']
        min           = [int]$row['min']
        tipo          = [string]$row['tipo']
    }
}

# Determinar ultimo diario (Full D), ultimo sabado (Full D >25 bases), y ultimo diferencial (I)
$fechasOrdenadas = @($sqlByDate.Keys | Sort-Object -Descending)

$ultimoDiario = $null; $ultimoDiarioFecha = $null
$ultimoSabado = $null; $ultimoSabadoFecha = $null
$ultimoDiff = $null; $ultimoDiffFecha = $null

foreach ($f in $fechasOrdenadas) {
    $items = $sqlByDate[$f]
    $itemsFull = @($items | Where-Object { $_.tipo -eq 'D' })
    $itemsDiff = @($items | Where-Object { $_.tipo -eq 'I' })
    
    if (-not $ultimoDiff -and $itemsDiff.Count -gt 0) {
        $ultimoDiff = $itemsDiff
        $ultimoDiffFecha = $f
    }
    if (-not $ultimoDiario -and $itemsFull.Count -gt 0 -and $itemsFull.Count -le 25) {
        $ultimoDiario = $itemsFull
        $ultimoDiarioFecha = $f
    }
    if (-not $ultimoSabado -and $itemsFull.Count -gt 25) {
        $ultimoSabado = $itemsFull
        $ultimoSabadoFecha = $f
    }
    if ($ultimoDiario -and $ultimoSabado -and $ultimoDiff) { break }
}

if (-not $ultimoDiario -and $fechasOrdenadas.Count -gt 0) {
    $fullItems = @($sqlByDate[$fechasOrdenadas[0]] | Where-Object { $_.tipo -eq 'D' })
    if ($fullItems.Count -gt 0) {
        $ultimoDiarioFecha = $fechasOrdenadas[0]
        $ultimoDiario = $fullItems
    }
}

# Construir Full Diario (tipo D)
$diffObj = @{ status = "NOT_FOUND"; fecha = ""; tipo = ""; bases = 0; tamano_gb = 0; comprimido_gb = 0; inicio = ""; fin = ""; max_min = 0; nota = "" }
$diffFiles = @()
if ($ultimoDiario) {
    $totalTam = ($ultimoDiario | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
    $totalComp = ($ultimoDiario | ForEach-Object { $_.comprimido_gb } | Measure-Object -Sum).Sum
    $maxMin = ($ultimoDiario | ForEach-Object { $_.min } | Measure-Object -Maximum).Maximum
    $sorted = @($ultimoDiario | Sort-Object { $_.inicio })
    $inicioMin = $sorted[0].inicio
    $finMax = ($ultimoDiario | Sort-Object { $_.fin } -Descending | Select-Object -First 1).fin
    
    $diffObj = @{
        status        = "COMPLETED"
        fecha         = $ultimoDiarioFecha
        tipo          = "D"
        bases         = $ultimoDiario.Count
        tamano_gb     = [math]::Round($totalTam, 2)
        comprimido_gb = [math]::Round($totalComp, 2)
        inicio        = $inicioMin
        fin           = $finMax
        max_min       = $maxMin
        nota          = "Full diario"
    }
    $diffFiles = $sorted
}

# Construir Diferencial real (tipo I)
$diffRealObj = @{ status = "NOT_FOUND"; fecha = ""; tipo = ""; bases = 0; tamano_gb = 0; comprimido_gb = 0; inicio = ""; fin = ""; max_min = 0; nota = "" }
$diffRealFiles = @()
if ($ultimoDiff) {
    $totalTam = ($ultimoDiff | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
    $totalComp = ($ultimoDiff | ForEach-Object { $_.comprimido_gb } | Measure-Object -Sum).Sum
    $maxMin = ($ultimoDiff | ForEach-Object { $_.min } | Measure-Object -Maximum).Maximum
    $sorted = @($ultimoDiff | Sort-Object { $_.inicio })
    $inicioMin = $sorted[0].inicio
    $finMax = ($ultimoDiff | Sort-Object { $_.fin } -Descending | Select-Object -First 1).fin
    
    $diffRealObj = @{
        status        = "COMPLETED"
        fecha         = $ultimoDiffFecha
        tipo          = "I"
        bases         = $ultimoDiff.Count
        tamano_gb     = [math]::Round($totalTam, 2)
        comprimido_gb = [math]::Round($totalComp, 2)
        inicio        = $inicioMin
        fin           = $finMax
        max_min       = $maxMin
        nota          = "Diferencial diario"
    }
    $diffRealFiles = $sorted
}

# Construir full sabado
$fullLocalObj = @{ status = "NOT_FOUND"; fecha = ""; tipo = ""; bases = 0; tamano_gb = 0; comprimido_gb = 0; inicio = ""; fin = ""; max_min = 0; nota = "Solo sabados" }
$fullLocalFiles = @()
if ($ultimoSabado) {
    $totalTam = ($ultimoSabado | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
    $totalComp = ($ultimoSabado | ForEach-Object { $_.comprimido_gb } | Measure-Object -Sum).Sum
    $maxMin = ($ultimoSabado | ForEach-Object { $_.min } | Measure-Object -Maximum).Maximum
    $sorted = @($ultimoSabado | Sort-Object { $_.inicio })
    $inicioMin = $sorted[0].inicio
    $finMax = ($ultimoSabado | Sort-Object { $_.fin } -Descending | Select-Object -First 1).fin
    $tipo = $ultimoSabado[0].tipo
    
    $fullLocalObj = @{
        status        = "COMPLETED"
        fecha         = $ultimoSabadoFecha
        tipo          = $tipo
        bases         = $ultimoSabado.Count
        tamano_gb     = [math]::Round($totalTam, 2)
        comprimido_gb = [math]::Round($totalComp, 2)
        inicio        = $inicioMin
        fin           = $finMax
        max_min       = $maxMin
        nota          = "Sabado - Full semanal $($ultimoSabado.Count) bases"
    }
    $fullLocalFiles = $sorted
}

# Historial MSSQL
$historialMssql = @()
foreach ($f in $fechasOrdenadas) {
    $items = $sqlByDate[$f]
    $totalTam = ($items | ForEach-Object { $_.tamano_gb } | Measure-Object -Sum).Sum
    $totalComp = ($items | ForEach-Object { $_.comprimido_gb } | Measure-Object -Sum).Sum
    $sorted = @($items | Sort-Object { $_.inicio })
    $inicioMin = $sorted[0].inicio
    $finMax = ($items | Sort-Object { $_.fin } -Descending | Select-Object -First 1).fin
    
    if ($items.Count -gt 25) {
        $historialMssql += @{
            fecha = $f
            full_local = @{
                status = "COMPLETED"; tipo = $items[0].tipo; bases = $items.Count
                tamano_gb = [math]::Round($totalTam, 2); comprimido_gb = [math]::Round($totalComp, 2)
                inicio = $inicioMin; fin = $finMax; nota = "Sabado - Full semanal $($items.Count) bases"
            }
            diff = @{ status = "NOT_FOUND"; bases = 0; tamano_gb = 0; comprimido_gb = 0; inicio = ""; fin = "" }
        }
    } else {
        $historialMssql += @{
            fecha = $f
            full_local = @{ status = "NOT_FOUND"; bases = 0; tamano_gb = 0; comprimido_gb = 0; inicio = ""; fin = ""; nota = "Solo sabados" }
            diff = @{
                status = "COMPLETED"; tipo = $items[0].tipo; bases = $items.Count
                tamano_gb = [math]::Round($totalTam, 2); comprimido_gb = [math]::Round($totalComp, 2)
                inicio = $inicioMin; fin = $finMax
            }
        }
    }
}

# ============================================================
# Generar JSON
# ============================================================
Write-Host "Generando JSON..."

$resultado = [ordered]@{
    fecha     = $hoy
    timestamp = $timestamp
    database  = "PROD"
    host      = "${oraHost}:${oraPort}"
    
    rman_full       = $rmanFullObj
    rman_full_files = $rmanFullFilesOut
    rman_arch       = $rmanArchObj
    rman_arch_files = $rmanArchFilesOut
    
    expdp       = $expdpObj
    expdp_files = $expdpFilesHoy
    
    historial = $historialOracle
    
    mssql = [ordered]@{
        server    = "SV109TIC03-P1"
        host      = $sqlServer
        database  = "IMPULSA"
        instancia = "BD SQL PROD"
        
        diferencial       = $diffObj
        full_local        = $fullLocalObj
        full_local_files  = $fullLocalFiles
        diferencial_files = $diffFiles
        diferencial_real       = $diffRealObj
        diferencial_real_files = $diffRealFiles
        
        historial_mssql = $historialMssql
    }
}

$json = $resultado | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.Encoding]::UTF8)

Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] backup_status.json actualizado correctamente"
Write-Host "  Oracle RMAN Full: $($rmanFullObj.status) ($($rmanFullFilesList.Count) piezas) | Arch: $($rmanArchObj.status)"
Write-Host "  SQL Server Diario: $($diffObj.status) ($($diffObj.bases) bases) | Sabado: $($fullLocalObj.status) ($($fullLocalObj.bases) bases)"
