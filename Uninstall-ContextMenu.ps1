<#
    Uninstall-ContextMenu.ps1
    Убирает оба пункта контекстного меню, добавленных Install-ContextMenu.ps1.
    Ничего не делает, если они не установлены — ошибок не будет.
#>

$ErrorActionPreference = "SilentlyContinue"

Remove-Item -Path "HKCU:\Software\Classes\*\shell\ConvertToMD" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\ConvertToMD_OCR" -Recurse -Force

Write-Host "Пункты контекстного меню удалены (если были установлены)." -ForegroundColor Green
Write-Host "Если не пропали сразу — перезапустите проводник: Stop-Process -Name explorer -Force"
