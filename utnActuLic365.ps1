# Nombre del Script: utnActuLic365.ps1
# Version: v1.3
# Autor: Lic. Ricardo MONLA
# Descripcion: Script para ...

# Verificar si puede ejecutar el modulo
Get-Command Set-MgBetaUserLicense

# Obtener la licencia
$licencia = Get-MgBetaSubscribedSku -All | Where-Object SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT'

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
    removeLicenses = @()
}

# Iterar sobre cada usuario y asignar licencia
foreach ($usr in $listaUSRs) {
    $uri = "https://graph.microsoft.com/v1.0/users/$usr/assignLicense"
    Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json -Depth 3) -ContentType "application/json"
    Write-Host "Licencia asignada a $usr"
}
