# ============================================================================
# monitoreo_server.ps1
# Servidor HTTP que sirve la pagina de monitoreo y consulta Oracle via SQLcl
# Puerto: 8080 -> http://localhost:8080
# ============================================================================
param(
    [int]$Port = 8080,
    [string]$SqlclPath = "sql",
    [string]$ConnectionName = "MCP_ORACLE",
    [int]$RefreshMinutes = 5
)

$WorkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$HtmlFile = Join-Path $WorkDir "monitoreo_live.html"
$SqlFile = Join-Path $WorkDir "monitoreo_query.sql"
$JsonCache = Join-Path $WorkDir "backup_status.json"
$JsonCacheLucia = Join-Path $WorkDir "backup_status_lucia.json"
$RefreshScript = Join-Path $WorkDir "refresh_data.ps1"
$RefreshScriptLucia = Join-Path $WorkDir "refresh_data_lucia.ps1"
$LastQuery = [DateTime]::MinValue
$LastQueryLucia = [DateTime]::MinValue

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$ts] $msg" -ForegroundColor Cyan
}

function Run-ScriptBackground($scriptPath, $timeoutSec = 90) {
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`"" -WorkingDirectory $WorkDir -WindowStyle Hidden -PassThru
    # No esperar - el script actualiza el JSON en disco
    # Guardar PID para controlar procesos huerfanos
    return $proc
}

$script:RefreshProcSI = $null
$script:RefreshProcSL = $null

function Get-BackupStatus {
    $now = Get-Date
    $cacheExists = Test-Path $JsonCache
    $needsRefresh = $false
    
    if ($cacheExists) {
        $fileAge = ($now - (Get-Item $JsonCache).LastWriteTime).TotalMinutes
        if ($fileAge -ge $RefreshMinutes) { $needsRefresh = $true }
    } else {
        $needsRefresh = $true
    }
    
    # Lanzar refresh en background si no hay uno corriendo
    if ($needsRefresh -and (Test-Path $RefreshScript)) {
        $running = $script:RefreshProcSI -and -not $script:RefreshProcSI.HasExited
        if (-not $running) {
            Write-Log "Ejecutando refresh_data.ps1 (background)..."
            $script:RefreshProcSI = Run-ScriptBackground $RefreshScript
        }
    }

    # Siempre retornar cache inmediatamente (aunque sea viejo)
    if ($cacheExists) {
        return Get-Content $JsonCache -Raw -Encoding UTF8
    }
    return '{"error":"No hay datos aun. Primera ejecucion en progreso..."}'
}

function Get-BackupStatusLucia {
    $now = Get-Date
    $cacheExists = Test-Path $JsonCacheLucia
    $needsRefresh = $false
    
    if ($cacheExists) {
        $fileAge = ($now - (Get-Item $JsonCacheLucia).LastWriteTime).TotalMinutes
        if ($fileAge -ge $RefreshMinutes) { $needsRefresh = $true }
    } else {
        $needsRefresh = $true
    }
    
    # Lanzar refresh en background si no hay uno corriendo
    if ($needsRefresh -and (Test-Path $RefreshScriptLucia)) {
        $running = $script:RefreshProcSL -and -not $script:RefreshProcSL.HasExited
        if (-not $running) {
            Write-Log "Ejecutando refresh_data_lucia.ps1 (background)..."
            $script:RefreshProcSL = Run-ScriptBackground $RefreshScriptLucia
        }
    }

    # Siempre retornar cache inmediatamente
    if ($cacheExists) {
        return Get-Content $JsonCacheLucia -Raw -Encoding UTF8
    }
    return '{"error":"No hay datos Santa Lucia aun. Primera ejecucion en progreso..."}'
}

# Generar datos iniciales usando SQLcl MCP (ya que la conexion MCP esta activa)
function Get-BackupStatusFromMCP {
    # Este metodo es llamado cuando SQLcl standalone no esta disponible
    # Retorna el JSON del cache si existe
    if (Test-Path $JsonCache) {
        return Get-Content $JsonCache -Raw -Encoding UTF8
    }
    return $null
}

# ============================================================================
# SERVIDOR HTTP
# ============================================================================
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/")

try {
    $listener.Start()
}
catch {
    Write-Host "ERROR: No se pudo iniciar el servidor en puerto $Port" -ForegroundColor Red
    Write-Host "Intente ejecutar como Administrador o cambiar el puerto" -ForegroundColor Yellow
    exit 1
}

Write-Host ""
Write-Host "======================================================" -ForegroundColor Green
Write-Host "  MONITOREO DE BACKUPS - ORACLE PROD" -ForegroundColor Green
Write-Host "  Servidor activo en: http://localhost:$Port" -ForegroundColor Green
Write-Host "  Auto-refresh cada: $RefreshMinutes minutos" -ForegroundColor Green
Write-Host "  Presione Ctrl+C para detener" -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Green
Write-Host ""

while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $path = $request.Url.LocalPath

        Write-Log "$($request.HttpMethod) $path"

        switch ($path) {
            "/api/status" {
                # Endpoint API: retorna JSON con estado de backups
                $json = Get-BackupStatus
                if (-not $json) { $json = '{"error":"Sin datos"}' }
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            "/api/refresh" {
                # Forzar reconsulta
                $script:LastQuery = [DateTime]::MinValue
                $json = Get-BackupStatus
                if (-not $json) { $json = '{"error":"Sin datos"}' }
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            "/api/status-lucia" {
                $json = Get-BackupStatusLucia
                if (-not $json) { $json = '{"error":"Sin datos Santa Lucia"}' }
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            "/api/refresh-lucia" {
                $script:LastQueryLucia = [DateTime]::MinValue
                $json = Get-BackupStatusLucia
                if (-not $json) { $json = '{"error":"Sin datos Santa Lucia"}' }
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                $response.ContentType = "application/json; charset=utf-8"
                $response.Headers.Add("Access-Control-Allow-Origin", "*")
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            default {
                # Servir archivos JSON estaticos o pagina HTML
                $localFile = Join-Path $WorkDir ($path.TrimStart('/'))
                if ($path -match '\.json$' -and (Test-Path $localFile)) {
                    $json = Get-Content $localFile -Raw -Encoding UTF8
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
                    $response.ContentType = "application/json; charset=utf-8"
                    $response.Headers.Add("Access-Control-Allow-Origin", "*")
                }
                elseif ($path -eq '/' -or $path -eq '' -or $path -eq '/index.html') {
                    if (Test-Path $HtmlFile) {
                        $html = Get-Content $HtmlFile -Raw -Encoding UTF8
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                        $response.ContentType = "text/html; charset=utf-8"
                    } else {
                        $msg = "Archivo $HtmlFile no encontrado"
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($msg)
                        $response.ContentType = "text/plain; charset=utf-8"
                        $response.StatusCode = 404
                    }
                }
                else {
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes('{"error":"Not found"}')
                    $response.ContentType = "application/json; charset=utf-8"
                    $response.StatusCode = 404
                }
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
        }
        $response.OutputStream.Close()
    }
    catch [System.Net.HttpListenerException] {
        break
    }
    catch {
        Write-Log "Error: $_"
    }
}

$listener.Stop()
Write-Log "Servidor detenido"
