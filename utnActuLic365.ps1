# Nombre del Script: utnActuLic365.ps1
# Version: v1.4
# Autor: Lic. Ricardo MONLA
# Descripcion: Script para ...


# --- Configuración Inicial y Conexión a Microsoft Graph ---

# 1. Verificar e Instalar el módulo Microsoft.Graph si no está presente
Write-Host "Verificando la instalación del módulo Microsoft.Graph..."
try {
    Import-Module Microsoft.Graph -ErrorAction Stop
    Write-Host "Módulo Microsoft.Graph ya instalado y cargado."
} catch {
    Write-Host "Módulo Microsoft.Graph no encontrado. Intentando instalar..."
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
        Import-Module Microsoft.Graph -ErrorAction Stop
        Write-Host "Módulo Microsoft.Graph instalado y cargado exitosamente."
    } catch {
        Write-Error "Error grave: No se pudo instalar el módulo Microsoft.Graph. Por favor, instálalo manualmente o verifica tus permisos."
        Write-Error "Detalles del error: $($_.Exception.Message)"
        exit
    }
}

# 2. Conectar a Microsoft Graph
Write-Host "Conectando a Microsoft Graph..."
try {
    # Intenta obtener una conexión existente
    $currentContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $currentContext) {
        # Si no hay conexión, solicita la conexión
        Write-Host "No se encontró una conexión activa a Microsoft Graph. Se abrirá una ventana para autenticar."
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All", "Directory.AccessAsUser.All", "Organization.Read.All"
        Write-Host "Conexión a Microsoft Graph establecida exitosamente."
    } else {
        Write-Host "Ya existe una conexión activa a Microsoft Graph."
        # Puedes añadir una comprobación para asegurar que los scopes necesarios estén presentes
        # Si no lo están, podrías pedir reconectar o avisar al usuario.
    }
} catch {
    Write-Error "No se pudo conectar a Microsoft Graph. Asegúrate de tener los permisos necesarios y vuelve a intentarlo."
    Write-Error "Detalles del error: $($_.Exception.Message)"
    exit
}


# --- PARTE 1: SELECCIÓN DEL ARCHIVO CSV Y LECTURA DE USUARIOS ---

Write-Host "`n--- Selección de Usuarios ---"
Write-Host "Por favor, selecciona el archivo CSV que contiene la lista de usuarios."

# Muestra un cuadro de diálogo para seleccionar un archivo CSV
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = [Environment]::GetFolderPath("Desktop")
$OpenFileDialog.filter = "Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*"
$OpenFileDialog.title = "Selecciona el archivo CSV con la lista de usuarios"
$OpenFileDialog.ShowHelp = $true

if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $csvFilePath = $OpenFileDialog.FileName
    Write-Host "Archivo CSV seleccionado: $csvFilePath"

    # Lee el archivo CSV y extrae las direcciones de correo electrónico
    # Asume que la columna con los emails se llama 'UserPrincipalName'.
    # Si tu CSV tiene otro nombre de columna para los emails, cámbialo aquí.
    try {
        $listaUSRs = (Import-Csv -Path $csvFilePath | Select-Object -ExpandProperty UserPrincipalName)
        if ($listaUSRs.Count -eq 0) {
            Write-Warning "El archivo CSV está vacío o no contiene la columna 'UserPrincipalName'."
            exit
        }
        Write-Host "Se encontraron $($listaUSRs.Count) usuarios en el CSV."
    } catch {
        Write-Error "Error al leer el archivo CSV o la columna 'UserPrincipalName' no se encontró. Detalles: $($_.Exception.Message)"
        exit
    }
} else {
    Write-Warning "No se seleccionó ningún archivo CSV. Saliendo del script."
    exit
}

# --- PARTE 2: PREPARACIÓN DE LA LICENCIA ---

Write-Host "`n--- Preparación de la Licencia ---"
# Obtener la licencia STANDARDWOFFPACK_STUDENT
Write-Host "Obteniendo detalles de la licencia STANDARDWOFFPACK_STUDENT..."
$licencia = Get-MgBetaSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT'
if (-not $licencia) {
    Write-Error "No se encontró la licencia 'STANDARDWOFFPACK_STUDENT' en tu tenant. Verifica el nombre o si está disponible."
    exit
}
Write-Host "Licencia STANDARDWOFFPACK_STUDENT encontrada. SkuId: $($licencia.SkuId)"

# Obtener los ServicePlans de la licencia que queremos asignar
$targetServicePlans = $licencia.ServicePlans | Select-Object -ExpandProperty ServicePlanId
Write-Host "La licencia STANDARDWOFFPACK_STUDENT contiene $($targetServicePlans.Count) Service Plans."

# --- PARTE 3: ASIGNACIÓN DE LICENCIAS CON DETECCIÓN DE CONFLICTOS ---

Write-Host "`n--- Iniciando Asignación de Licencias ---"

foreach ($usr in $listaUSRs) {
    Write-Host "`nProcesando usuario: $usr"
    $uri = "https://graph.microsoft.com/v1.0/users/$usr/assignLicense"

    # 1. Obtener las licencias actuales del usuario
    Write-Host "  Obteniendo licencias actuales para $usr..."
    $userLicenses = Get-MgUserLicenseDetail -UserId $usr -ErrorAction SilentlyContinue
    if (-not $userLicenses) {
        Write-Warning "  No se pudieron obtener los detalles de licencia para $usr. Saltando a este usuario."
        continue
    }

    $currentUserServicePlans = @()
    foreach ($userLicense in $userLicenses) {
        $currentUserServicePlans += $userLicense.ServicePlans | Select-Object -ExpandProperty ServicePlanId
    }
    Write-Host "  $usr tiene $($currentUserServicePlans.Count) Service Plans activos."

    # 2. Detectar Service Plans en conflicto
    $disabledPlans = @()
    $conflictsFound = $false
    foreach ($targetPlanId in $targetServicePlans) {
        if ($currentUserServicePlans -contains $targetPlanId) {
            $matchingPlanName = ($licencia.ServicePlans | Where-Object ServicePlanId -eq $targetPlanId | Select-Object -ExpandProperty ServicePlanName)
            Write-Warning "  Conflicto detectado: El Service Plan '$matchingPlanName' ($targetPlanId) ya existe para $usr. Se agregará a la lista de planes a deshabilitar."
            $disabledPlans += $targetPlanId
            $conflictsFound = $true
        }
    }

    if (-not $conflictsFound) {
        Write-Host "  No se detectaron conflictos de Service Plans para $usr."
    }

    # 3. Construir el cuerpo de la petición con los planes a deshabilitar si existen conflictos
    $body = @{
        addLicenses    = @(@{
            skuId        = $licencia.SkuId
            disabledPlans = $disabledPlans
        })
        removeLicenses = @()
    }

    # 4. Asignar la licencia
    try {
        Write-Host "  Asignando licencia STANDARDWOFFPACK_STUDENT a $usr..."
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json -Depth 3) -ContentType "application/json"
        if ($disabledPlans.Count -gt 0) {
            Write-Host "  Licencia asignada exitosamente a $usr (con $($disabledPlans.Count) planes deshabilitados)."
        } else {
            Write-Host "  Licencia asignada exitosamente a $usr (sin planes deshabilitados)."
        }
    } catch {
        Write-Error "  Error al asignar licencia a $usr. Detalles: $($_.Exception.Message)"
    }
}

Write-Host "`n--- Proceso de asignación de licencias completado. ---"

# Desconectar de Microsoft Graph (opcional)
# Write-Host "Desconectando de Microsoft Graph..."
# Disconnect-MgGraph
# Write-Host "Desconexión completa."