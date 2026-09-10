<#
    Uninstall-ContextMenu.ps1
    Убирает оба пункта контекстного меню, добавленных Install-ContextMenu.ps1.

    Удаляет и текущую регистрацию (по каждому формату документов), и старую
    регистрацию первой версии (пункт на всех файлах, ключ "*"), поэтому
    подходит и для обновления со старой версии.

    Ничего не делает, если пункты не установлены — ошибок не будет.
    Файлы документов и результаты конвертации не затрагиваются: скрипт
    работает только с реестром.
#>

$ErrorActionPreference = "SilentlyContinue"

$menuExtensions = @(
    ".pdf", ".docx", ".xlsx", ".xls", ".pptx",
    ".html", ".htm", ".csv", ".txt", ".epub",
    ".doc", ".ppt"
)

$verbConvert = "ConvertToMD"
$verbOcr     = "ConvertToMD_OCR"

$removed = 0

function Remove-RegistryKeyIfExists {
    param([string]$LiteralKeyPath)
    # -LiteralPath обязателен: в старом пути есть символ "*", который иначе
    # был бы истолкован как подстановочный знак и задел бы соседние ключи
    if (Test-Path -LiteralPath $LiteralKeyPath) {
        Remove-Item -LiteralPath $LiteralKeyPath -Recurse -Force
        return $true
    }
    return $false
}

# Регистрация первой версии — пункт показывался для всех файлов
if (Remove-RegistryKeyIfExists "HKCU:\Software\Classes\*\shell\$verbConvert") { $removed++ }

# Текущая регистрация — по каждому формату документов
foreach ($extension in $menuExtensions) {
    if (Remove-RegistryKeyIfExists "HKCU:\Software\Classes\SystemFileAssociations\$extension\shell\$verbConvert") { $removed++ }
}

# Пункт с OCR
if (Remove-RegistryKeyIfExists "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\$verbOcr") { $removed++ }

if ($removed -gt 0) {
    Write-Host "Пункты контекстного меню удалены (записей в реестре: $removed)." -ForegroundColor Green
} else {
    Write-Host "Пункты контекстного меню не были установлены — удалять нечего." -ForegroundColor Yellow
}

Write-Host "Если пункты не пропали сразу — перезапустите проводник:"
Write-Host "  Stop-Process -Name explorer -Force"
