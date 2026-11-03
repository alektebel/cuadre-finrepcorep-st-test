<#
.SYNOPSIS
  Convierte el libro generado (.xlsx) en .xlsm con el modulo VBA de carga y los botones.

.DESCRIPTION
  Requiere Windows + Excel y, una sola vez, activar en Excel:
  Archivo > Opciones > Centro de confianza > Configuracion > Configuracion de macros >
  "Confiar en el acceso al modelo de objetos de proyectos de VBA".

.EXAMPLE
  cuadre excel --ejercicio config/exercises/st2025_2024-12.yaml --salida out/cuadre.xlsx
  powershell -ExecutionPolicy Bypass -File tools/build_xlsm.ps1 -Xlsx out/cuadre.xlsx
#>
param(
    [Parameter(Mandatory = $true)][string]$Xlsx,
    [string]$Bas = (Join-Path $PSScriptRoot "..\src\cuadre\excel_book\vba\Cuadre.bas"),
    [string]$Out
)
$ErrorActionPreference = "Stop"
$xlsxPath = (Resolve-Path $Xlsx).Path
$basPath = (Resolve-Path $Bas).Path
if (-not $Out) { $Out = [System.IO.Path]::ChangeExtension($xlsxPath, ".xlsm") }
$outPath = [System.IO.Path]::GetFullPath($Out)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($xlsxPath)
    try {
        $null = $wb.VBProject.VBComponents.Import($basPath)
    } catch {
        throw "No se pudo importar el modulo VBA. Activa 'Confiar en el acceso al modelo de objetos de proyectos de VBA'. Detalle: $_"
    }
    $wb.SaveAs($outPath, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled
    $excel.Run("'" + $wb.Name + "'!CrearBotones")
    $wb.Save()
    $wb.Close($false)
    Write-Host "Libro con macros: $outPath"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
