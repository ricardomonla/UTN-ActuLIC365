# Nombre del Script: utnActuLic365.ps1
# Version: v2.0.0
# Autor: Lic. Ricardo MONLA (Modificado por IA)
# Descripcion: Script para asignar licencias de Microsoft 365 a usuarios con menú interactivo, carga CSV y logging.

# --- Global Variables ---
$global:licencia = $null
$global:listaUSRs = @()
$global:logFolderPath = Join-Path $PSScriptRoot "Logs"
$global:successLogFile = Join-Path $global:logFolderPath "licencias_exitosas_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$global:errorLogFile = Join-Path $global:logFolderPath "licencias_errores_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Create the Logs directory if it doesn't exist
if (-not (Test-Path $global:logFolderPath -PathType Container)) {
    New-Item -Path $global:logFolderPath -ItemType Directory | Out-Null
}

# --- Functions ---

function Show-Menu {
    Clear-Host
    Write-Host "==============================================="
    Write-Host "           Gestor de Licencias M365            "
    Write-Host "==============================================="
    Write-Host "1. Obtener información de la Licencia"
    Write-Host "2. Cargar lista de usuarios desde archivo CSV"
    Write-Host "3. Asignar Licencias a usuarios cargados"
    Write-Host "4. Salir"
    Write-Host "-----------------------------------------------"

    if ($global:licencia) {
        Write-Host "Licencia actual: $($global:licencia.SkuPartNumber) (ID: $($global:licencia.SkuId))" -ForegroundColor DarkCyan
    } else {
        Write-Host "No se ha obtenido información de licencia." -ForegroundColor DarkYellow
    }
    if ($global:listaUSRs.Count -gt 0) {
        Write-Host "Usuarios cargados: $($global:listaUSRs.Count) (Listos para asignar)" -ForegroundColor DarkCyan
    } else {
        Write-Host "No hay usuarios cargados." -ForegroundColor DarkYellow
    }
    Write-Host "-----------------------------------------------"
}

function Get-LicenseInfo {
    Write-Host "`n--- Obteniendo información de la licencia ---" -ForegroundColor Yellow

    # Conectar a Microsoft Graph si no está conectado
    if (-not (Get-MgContext)) {
        Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Green
        Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All" -NoWelcome
    } else {
        Write-Host "Ya conectado a Microsoft Graph." -ForegroundColor Green
    }

    # Obtener la licencia
    # Puedes ver las licencias disponibles con Get-MgBetaSubscribedSku -All | Select SkuPartNumber, SkuId
    try {
        $global:licencia = Get-MgBetaSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT'

        if (-not $global:licencia) {
            Write-Warning "No se encontró la licencia con SkuPartNumber 'STANDARDWOFFPACK_STUDENT'. Por favor, verifica el nombre."
            Add-Content -Path $global:errorLogFile -Value "$(Get-Date) - ERROR: No se encontró la licencia 'STANDARDWOFFPACK_STUDENT'."
        } else {
            Write-Host "Licencia encontrada: $($global:licencia.SkuPartNumber) con SkuId: $($global:licencia.SkuId)" -ForegroundColor Green
            Add-Content -Path $global:successLogFile -Value "$(Get-Date) - INFO: Licencia 'STANDARDWOFFPACK_STUDENT' (ID: $($global:licencia.SkuId)) obtenida."
        }
    } catch {
        Write-Error "Error al intentar obtener la licencia: $($_.Exception.Message)"
        Add-Content -Path $global:errorLogFile -Value "$(Get-Date) - ERROR: Error al obtener la licencia: $($_.Exception.Message)"
    }
    Pause-Script
}

function Load-UsersFromCsv {
    Write-Host "`n--- Cargando usuarios desde archivo CSV ---" -ForegroundColor Yellow

    # Abrir ventana para seleccionar archivo CSV
    $csvFilePath = Show-FileDialog -Title "Seleccionar archivo CSV de usuarios" -Filter "Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*"

    if (-not $csvFilePath) {
        Write-Warning "No se seleccionó ningún archivo CSV."
        return
    }

    if (-not (Test-Path $csvFilePath)) {
        Write-Error "El archivo '$csvFilePath' no existe."
        return
    }

    try {
        $csvContent = Import-Csv -Path $csvFilePath -Header "UserPrincipalName"
        $global:listaUSRs = $csvContent.UserPrincipalName | Where-Object { $_ -match '.@.' } # Basic validation for UPN format
        Write-Host "Se cargaron $($global:listaUSRs.Count) usuarios desde '$csvFilePath'." -ForegroundColor Green
        Add-Content -Path $global:successLogFile -Value "$(Get-Date) - INFO: Se cargaron $($global:listaUSRs.Count) usuarios desde '$csvFilePath'."
    } catch {
        Write-Error "Error al leer el archivo CSV '$csvFilePath': $($_.Exception.Message)"
        Add-Content -Path $global:errorLogFile -Value "$(Get-Date) - ERROR: Error al leer el archivo CSV '$csvFilePath': $($_.Exception.Message)"
    }
    Pause-Script
}

# Helper function for file dialog (requires System.Windows.Forms)
function Show-FileDialog {
    param(
        [string]$Title = "Seleccionar archivo",
        [string]$Filter = "Todos los archivos (*.*)|*.*"
    )

    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = $Title
    $fileDialog.Filter = $Filter
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments') # Start in MyDocuments

    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    return $null
}

function Assign-Licenses {
    Write-Host "`n--- Asignando licencias ---" -ForegroundColor Yellow

    if (-not $global:licencia) {
        Write-Warning "Primero debe obtener la información de la licencia (Opción 1)."
        Pause-Script
        return
    }
    if ($global:listaUSRs.Count -eq 0) {
        Write-Warning "Primero debe cargar la lista de usuarios desde un CSV (Opción 2)."
        Pause-Script
        return
    }

    $assignedCount = 0
    $errorCount = 0
    $successUsers = @()
    $errorUsers = @()

    $body = @{
        addLicenses    = @(@{ skuId = $global:licencia.SkuId })
        removeLicenses = @()
    }

    Write-Host "Iniciando asignación de licencias para $($global:listaUSRs.Count) usuarios..." -ForegroundColor Cyan

    foreach ($usr in $global:listaUSRs) {
        Write-Host "Asignando licencia a ${usr}..." -NoNewline

        try {
            $uri = "https://graph.microsoft.com/v1.0/users/$usr/assignLicense"
            Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json -Depth 3) -ContentType "application/json" -ErrorAction Stop

            Write-Host " [ÉXITO]" -ForegroundColor Green
            $assignedCount++
            $successUsers += $usr
            Add-Content -Path $global:successLogFile -Value "$(Get-Date) - ÉXITO: Licencia asignada a ${usr}."
        } catch {
            Write-Host " [ERROR]" -ForegroundColor Red
            $errorCount++
            $errorMessage = "Error al asignar licencia a ${usr}: $($_.Exception.Message)"
            Write-Error $errorMessage
            $errorUsers += $usr
            Add-Content -Path $global:errorLogFile -Value "$(Get-Date) - ERROR: ${errorMessage}"
        }
    }

    Write-Host "`n--- Resumen de asignación de licencias ---" -ForegroundColor Yellow
    Write-Host "Total de usuarios procesados: $($global:listaUSRs.Count)"
    Write-Host "Licencias asignadas exitosamente: $assignedCount" -ForegroundColor Green
    Write-Host "Errores en la asignación: $errorCount" -ForegroundColor Red

    if ($successUsers.Count -gt 0) {
        Write-Host "`nUsuarios con licencia asignada (registrados en '$($global:successLogFile)'):" -ForegroundColor Green
        $successUsers | ForEach-Object { Write-Host "  - $_" }
    }
    if ($errorUsers.Count -gt 0) {
        Write-Host "`nUsuarios con errores en la asignación (registrados en '$($global:errorLogFile)'):" -ForegroundColor Red
        $errorUsers | ForEach-Object { Write-Host "  - $_" }
    }

    Pause-Script
}

function Pause-Script {
    Write-Host "`nPresiona cualquier tecla para continuar..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
}

# --- Main Program Loop ---
while ($true) {
    Show-Menu
    $choice = Read-Host "Ingresa tu opción [1-4]"

    switch ($choice) {
        "1" { Get-LicenseInfo }
        "2" { Load-UsersFromCsv }
        "3" { Assign-Licenses }
        "4" {
            Write-Host "Desconectando de Microsoft Graph y saliendo..." -ForegroundColor DarkCyan
            Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue
            exit
        }
        default {
            Write-Warning "Opción no válida. Por favor, selecciona entre 1 y 4."
            Pause-Script
        }
    }
}