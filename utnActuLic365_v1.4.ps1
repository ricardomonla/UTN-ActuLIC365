# Nombre del Script: utnActuLic365.ps1
# Version: v1.4
# Autor: Lic. Ricardo MONLA
# Descripcion: Script para asignar licencias de Microsoft 365 a usuarios.

# Importar el módulo Microsoft.Graph si no está ya importado
# If (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
#     Install-Module -Name Microsoft.Graph -Scope CurrentUser
# }
# Import-Module Microsoft.Graph

# Conectar a Microsoft Graph
# Necesitas los siguientes permisos:
#   - Organization.Read.All (para Get-MgBetaSubscribedSku)
#   - User.ReadWrite.All (para asignar licencias a usuarios)
#   - Directory.Read.All (para listar usuarios si fuera necesario de otra forma)
Connect-MgGraph -Scopes "Organization.Read.All", "User.ReadWrite.All"

# Verificar si puede ejecutar el modulo (Esto es más para verificar la instalación del módulo, no la autenticación)
# Get-Command Set-MgBetaUserLicense # This command requires the 'Microsoft.Graph.Beta.Users' module.

# Obtener la licencia
# Asegúrate de que el SkuPartNumber sea correcto para tu tenant.
# Puedes ver las licencias disponibles con Get-MgBetaSubscribedSku -All | Select SkuPartNumber, SkuId
$licencia = Get-MgBetaSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT'

# Verificar si se encontró la licencia
if (-not $licencia) {
    Write-Error "No se encontró la licencia con SkuPartNumber 'STANDARDWOFFPACK_STUDENT'. Por favor, verifica el nombre."
    exit
}

Write-Host "Licencia encontrada: $($licencia.SkuPartNumber) con SkuId: $($licencia.SkuId)"

# Lista de usuarios (array de strings)
$listaUSRs = @(
    "aguilarguardia.santiago@frlr.utn.edu.ar"
    "dell.32731@frlr.utn.edu.ar"
    "dell.32732@frlr.utn.edu.ar"
    "dell.32733@frlr.utn.edu.ar"
    "dell.32734@frlr.utn.edu.ar"
    "salva.natalia@frlr.utn.edu.ar"
)

# Cuerpo de la petición (convertido a JSON al final)
$body = @{
    addLicenses    = @(@{ skuId = $licencia.SkuId })
    removeLicenses = @() # Si no necesitas remover ninguna licencia, déjalo vacío.
}

# Iterar sobre cada usuario y asignar licencia
foreach ($usr in $listaUSRs) {
    try {
        # Para asignar licencias a usuarios de forma más directa, puedes usar Set-MgUserLicense
        # Pero tu enfoque de Invoke-MgGraphRequest con assignLicense también es válido si prefieres la API directa.
        # Asegúrate de que el ID del usuario sea correcto. A veces el UPN (UserPrincipalName) no es el ID del objeto.
        # Para obtener el ID del usuario: Get-MgUser -UserId $usr | Select-Object Id

        # Usando Invoke-MgGraphRequest (tu método actual)
        # La URL para asignar licencias es /users/{id | userPrincipalName}/assignLicense
        # El ID o UserPrincipalName debe ser el del usuario en Azure AD.
        # Asumiendo que $usr es el UserPrincipalName.
        $uri = "https://graph.microsoft.com/v1.0/users/$usr/assignLicense"
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json -Depth 3) -ContentType "application/json"
        Write-Host "Licencia asignada a $usr" -ForegroundColor Green
    }
    catch {
        Write-Error "Error al asignar licencia a $usr: $($_.Exception.Message)"
        # Puedes añadir más lógica aquí para manejar errores específicos, como usuarios no encontrados.
    }
}

# Opcional: Desconectar de Microsoft Graph al finalizar
# Disconnect-MgGraph