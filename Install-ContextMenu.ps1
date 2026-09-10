<#
    Install-ContextMenu.ps1
    Добавляет в контекстное меню проводника два пункта:
      - "Конвертировать в MD"          — для поддерживаемых документов (без OCR)
      - "Конвертировать в MD (с OCR)"  — только для .pdf

    Работает через реестр текущего пользователя (HKEY_CURRENT_USER) —
    права администратора НЕ нужны, для других учётных записей на этом
    компьютере ничего не меняется.

    Пункт регистрируется только на форматах документов. В первой версии он
    регистрировался на "*", то есть показывался вообще для любого файла —
    включая архивы, картинки и программы, где конвертация заведомо
    закончилась бы сообщением "формат не поддерживается". Старая
    регистрация на "*" при установке удаляется.

    Поддержан мультивыбор: если выделить несколько файлов и нажать пункт
    меню, команда выполнится отдельно для каждого файла. Учтите, что для
    такого способа регистрации Windows ограничивает выделение примерно
    сотней файлов — для больших объёмов используйте папочный скрипт.

    Запуск: правой кнопкой по этому файлу -> "Выполнить с помощью PowerShell",
    либо в PowerShell: .\Install-ContextMenu.ps1
    После установки может понадобиться перезапустить проводник, чтобы меню
    обновилось.
#>

$ErrorActionPreference = "Stop"

# Папка, где лежит установщик — в ней же должны лежать рабочие скрипты.
# Реестр будет ссылаться именно на эти пути, поэтому после установки папку
# перемещать нельзя (иначе пункты меню перестанут работать).
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$convertScript    = Join-Path $scriptDir "ConvertToMD.ps1"
$convertOcrScript = Join-Path $scriptDir "ConvertToMD-OCR.ps1"
$commonScript     = Join-Path $scriptDir "ConvertToMD.Common.ps1"

foreach ($required in @($convertScript, $convertOcrScript, $commonScript)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host "Не найден файл: $required" -ForegroundColor Red
        Write-Host "Все скрипты должны лежать в одной папке с установщиком." -ForegroundColor Red
        exit 1
    }
}

# Расширения, для которых показывать пункт "Конвертировать в MD".
# .doc и .ppt включены намеренно: markitdown их не поддерживает, но скрипт
# покажет понятное сообщение с советом пересохранить в .docx / .pptx —
# это полезнее, чем отсутствие пункта меню без всяких объяснений.
$menuExtensions = @(
    ".pdf", ".docx", ".xlsx", ".xls", ".pptx",
    ".html", ".htm", ".csv", ".txt", ".epub",
    ".doc", ".ppt"
)

$verbConvert = "ConvertToMD"
$verbOcr     = "ConvertToMD_OCR"

function Remove-RegistryKeyIfExists {
    param([string]$LiteralKeyPath)
    # -LiteralPath обязателен: в старом пути есть символ "*", который иначе
    # был бы истолкован как подстановочный знак
    if (Test-Path -LiteralPath $LiteralKeyPath) {
        Remove-Item -LiteralPath $LiteralKeyPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Register-ContextMenuVerb {
    param(
        [string]$KeyPath,
        [string]$Title,
        [string]$TargetScript
    )
    New-Item -Path $KeyPath -Force | Out-Null
    Set-ItemProperty -LiteralPath $KeyPath -Name "(Default)" -Value $Title
    # Player — команда выполняется отдельно для каждого выделенного файла
    Set-ItemProperty -LiteralPath $KeyPath -Name "MultiSelectModel" -Value "Player"

    $commandKey = Join-Path $KeyPath "command"
    New-Item -Path $commandKey -Force | Out-Null
    $command = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$TargetScript`" `"%1`""
    Set-ItemProperty -LiteralPath $commandKey -Name "(Default)" -Value $command
}

# --- Убираем регистрацию первой версии (пункт на всех файлах) ---
Remove-RegistryKeyIfExists "HKCU:\Software\Classes\*\shell\$verbConvert"

# --- Пункт 1: "Конвертировать в MD" — по каждому формату документов ---
foreach ($extension in $menuExtensions) {
    $keyPath = "HKCU:\Software\Classes\SystemFileAssociations\$extension\shell\$verbConvert"
    Register-ContextMenuVerb -KeyPath $keyPath -Title "Конвертировать в MD" -TargetScript $convertScript
}

# --- Пункт 2: "Конвертировать в MD (с OCR)" — только для .pdf ---
$ocrKeyPath = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\$verbOcr"
Register-ContextMenuVerb -KeyPath $ocrKeyPath -Title "Конвертировать в MD (с OCR)" -TargetScript $convertOcrScript

Write-Host "Пункты меню добавлены:" -ForegroundColor Green
Write-Host "  - 'Конвертировать в MD' — для форматов: $($menuExtensions -join ', ')"
Write-Host "  - 'Конвертировать в MD (с OCR)' — только для .pdf"
Write-Host ""

# --- Проверка окружения ---
# Наличие команды python ещё не значит, что в ней есть нужный модуль:
# на компьютере может стоять несколько установок Python. Проверяем импортом.

function Test-PythonModule {
    param([string]$Module, [string[]]$Candidates)
    foreach ($candidate in $Candidates) {
        if (-not (Get-Command $candidate -ErrorAction SilentlyContinue)) { continue }
        try {
            $null = & $candidate -c "import $Module" 2>&1
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch { }
    }
    return $null
}

Write-Host "Проверка установленных компонентов:" -ForegroundColor Cyan

$foundMarkitdown = Test-PythonModule -Module "markitdown" -Candidates @("python", "python3")
if ($foundMarkitdown) {
    Write-Host "  markitdown — найден ($foundMarkitdown)" -ForegroundColor Green
} else {
    Write-Host "  markitdown — НЕ НАЙДЕН" -ForegroundColor Yellow
    Write-Host "    Установите:  python -m pip install --upgrade `"markitdown[pdf,docx,pptx,xlsx,xls]`""
}

$foundOcr = Test-PythonModule -Module "ocrmypdf" -Candidates @("python3", "python")
if ($foundOcr) {
    Write-Host "  ocrmypdf — найден ($foundOcr)" -ForegroundColor Green
} else {
    Write-Host "  ocrmypdf — НЕ НАЙДЕН (нужен только для пункта с OCR)" -ForegroundColor Yellow
    Write-Host "    Установите:  python -m pip install --upgrade ocrmypdf"
    Write-Host "    Дополнительно нужны Tesseract OCR (с пакетами rus и eng) и Ghostscript."
}

# Отдельно проверяем компоненты markitdown для форматов: базовой установки
# markitdown недостаточно, поддержка PDF, Word, Excel и PowerPoint вынесена
# в дополнительные наборы зависимостей
if ($foundMarkitdown) {
    $formatModules = [ordered]@{
        "PDF"        = "pdfminer"
        "Word"       = "mammoth"
        "Excel"      = "openpyxl"
        "PowerPoint" = "pptx"
    }
    $missing = @()
    foreach ($name in $formatModules.Keys) {
        $null = & $foundMarkitdown -c "import $($formatModules[$name])" 2>&1
        if ($LASTEXITCODE -ne 0) { $missing += $name }
    }
    if ($missing.Count -gt 0) {
        Write-Host "  Не хватает компонентов markitdown для форматов: $($missing -join ', ')" -ForegroundColor Yellow
        Write-Host "    Установите:  python -m pip install --upgrade `"markitdown[pdf,docx,pptx,xlsx,xls]`""
    } else {
        Write-Host "  Компоненты markitdown для PDF, Word, Excel и PowerPoint — на месте" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "Если пункты не появились сразу — перезапустите проводник:" -ForegroundColor Yellow
Write-Host "  Stop-Process -Name explorer -Force"
Write-Host "(он перезапустится автоматически), либо перезайдите в систему."
