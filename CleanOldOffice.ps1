# Script para limpiar restos de Office antiguos y permitir instalación de Office 365

# Desinstalar versiones antiguas de Office
$oldOfficeVersions = @(
    "Microsoft Office 2007",
    "Microsoft Office 2010",
    "Microsoft Office 2013",
    "Microsoft Office 2016",
    "Microsoft Office 2019"
)

$installed = Get-WmiObject -Query "SELECT * FROM Win32_Product" | Where-Object {
    $oldOfficeVersions -contains $_.Name
}

if ($installed) {
    foreach ($office in $installed) {
        Write-Host "Desinstalando: $($office.Name)"
        $office.Uninstall()
    }
}

# Eliminar carpetas residuales de Office
$folders = @(
    "$env:ProgramFiles\Microsoft Office",
    "$env:ProgramFiles (x86)\Microsoft Office",
    "$env:ProgramData\Microsoft\Office"
)

foreach ($folder in $folders) {
    if (Test-Path $folder) {
        Write-Host "Eliminando carpeta: $folder"
        Remove-Item -Path $folder -Recurse -Force
    }
}

# Eliminar claves de registro de Office
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
    "HKCU:\Software\Microsoft\Office"
)

foreach ($reg in $regPaths) {
    if (Test-Path $reg) {
        Write-Host "Eliminando clave de registro: $reg"
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "Limpieza completada. Reinicia el equipo antes de instalar Office 365."