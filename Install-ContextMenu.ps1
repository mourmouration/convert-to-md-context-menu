<#
    Install-ContextMenu.ps1
    Добавляет в контекстное меню проводника два пункта:
      - "Конвертировать в MD"          — для любого поддерживаемого файла (без OCR)
      - "Конвертировать в MD (с OCR)"  — только для .pdf

    Работает через реестр текущего пользователя (HKEY_CURRENT_USER) —
    права администратора НЕ нужны, ничего не меняется для других учётных
    записей на этом компьютере.

    Поддержан мультивыбор: если выделить несколько файлов и кликнуть один
    пункт меню, команда запустится отдельно для каждого файла.

    Запуск: правой кнопкой по этому файлу -> "Выполнить с помощью PowerShell",
    либо в PowerShell: .\Install-ContextMenu.ps1
    После установки может понадобиться перезапустить проводник (или просто
    перезайти в систему), чтобы меню обновилось.
#>

$ErrorActionPreference = "Stop"

# Папка, где лежит этот установщик — в ней же должны лежать ConvertToMD.ps1
# и ConvertToMD-OCR.ps1 (не переносите их отдельно от установщика).
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$convertScript    = Join-Path $scriptDir "ConvertToMD.ps1"
$convertOcrScript = Join-Path $scriptDir "ConvertToMD-OCR.ps1"

if (-not (Test-Path -LiteralPath $convertScript)) {
    Write-Host "Не найден файл: $convertScript" -ForegroundColor Red
    Write-Host "Положите ConvertToMD.ps1 рядом с этим установщиком и запустите снова." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $convertOcrScript)) {
    Write-Host "Не найден файл: $convertOcrScript" -ForegroundColor Red
    Write-Host "Положите ConvertToMD-OCR.ps1 рядом с этим установщиком и запустите снова." -ForegroundColor Red
    exit 1
}

# --- Пункт 1: "Конвертировать в MD" — для всех файлов (*) ---
$keyAll = "HKCU:\Software\Classes\*\shell\ConvertToMD"
New-Item -Path $keyAll -Force | Out-Null
Set-ItemProperty -Path $keyAll -Name "(Default)" -Value "Конвертировать в MD"
Set-ItemProperty -Path $keyAll -Name "MultiSelectModel" -Value "Player"

$keyAllCmd = "$keyAll\command"
New-Item -Path $keyAllCmd -Force | Out-Null
$commandAll = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$convertScript`" `"%1`""
Set-ItemProperty -Path $keyAllCmd -Name "(Default)" -Value $commandAll

# --- Пункт 2: "Конвертировать в MD (с OCR)" — только для .pdf ---
$keyPdf = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\ConvertToMD_OCR"
New-Item -Path $keyPdf -Force | Out-Null
Set-ItemProperty -Path $keyPdf -Name "(Default)" -Value "Конвертировать в MD (с OCR)"
Set-ItemProperty -Path $keyPdf -Name "MultiSelectModel" -Value "Player"

$keyPdfCmd = "$keyPdf\command"
New-Item -Path $keyPdfCmd -Force | Out-Null
$commandPdf = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$convertOcrScript`" `"%1`""
Set-ItemProperty -Path $keyPdfCmd -Name "(Default)" -Value $commandPdf

Write-Host "Готово. Пункты меню добавлены:" -ForegroundColor Green
Write-Host "  - 'Конвертировать в MD' — для всех поддерживаемых файлов"
Write-Host "  - 'Конвертировать в MD (с OCR)' — только для .pdf"
Write-Host ""
Write-Host "Если пункты не появились сразу — перезапустите проводник:" -ForegroundColor Yellow
Write-Host "  Stop-Process -Name explorer -Force"
Write-Host "(он перезапустится автоматически), либо перезайдите в систему."
