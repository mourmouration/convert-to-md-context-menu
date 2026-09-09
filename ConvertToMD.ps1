<#
    ConvertToMD.ps1
    Вызывается из контекстного меню проводника (пункт "Конвертировать в MD").
    Конвертирует ОДИН файл напрямую через markitdown, без OCR.
    Результат кладётся в подпапку "0.1. MD" рядом с исходным файлом —
    так же, как в папочных скриптах pdf_to_md_*.

    Поддерживаются: .pdf, .docx, .xls, .xlsx, .ppt, .pptx, .html, .htm,
    .csv, .txt, .epub. Для .pdf без текстового слоя (сканы) результат может
    получиться пустым или почти пустым — для таких файлов используйте пункт
    "Конвертировать в MD (с OCR)".
    .doc (Word 97-2003) markitdown не поддерживает — такой файл пропускается.

    На успехе скрипт молчит (файл просто появляется в "0.1. MD") — это
    важно при конвертации сразу нескольких выделенных файлов, чтобы не
    заваливать экран всплывающими окнами. Об ошибке или неподдерживаемом
    формате сообщает окном — там что-то реально пошло не так.

    ВАЖНО про $ErrorActionPreference: специально НЕ ставится в "Stop".
    markitdown при обычной успешной работе может писать диагностические
    сообщения в stderr — PowerShell способен воспринять это как
    завершающую ошибку и оборвать скрипт до создания .md, даже если
    конвертация прошла нормально. Поэтому вызов обёрнут в try/catch, а об
    успехе/неудаче судим не по коду возврата, а по тому, появился ли
    реально файл .md.
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Show-Message([string]$text, [string]$icon = "Warning") {
    [System.Windows.Forms.MessageBox]::Show($text, "Конвертация в MD", "OK", $icon) | Out-Null
}

function Test-CommandExists($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    Show-Message "Файл не найден:`n$Path"
    exit 1
}

if (Test-CommandExists "python") {
    $pythonForMd = "python"
} elseif (Test-CommandExists "python3") {
    $pythonForMd = "python3"
} else {
    Show-Message "Не найдена команда 'python' (и 'python3') в PATH.`nПроверьте установку Python / markitdown." "Error"
    exit 1
}

# Глушим безобидное предупреждение pydub про отсутствующий ffmpeg —
# оно не влияет на конвертацию не-аудио форматов.
$env:PYTHONWARNINGS = "ignore::RuntimeWarning:pydub.utils"

$directExtensions      = @(".pdf", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".html", ".htm", ".csv", ".txt", ".epub")
$unsupportedExtensions = @(".doc")

$file = Get-Item -LiteralPath $Path
$ext  = $file.Extension.ToLower()

if ($unsupportedExtensions -contains $ext) {
    Show-Message "Файл '$($file.Name)':`n.doc (Word 97-2003) не поддерживается markitdown.`nПересохраните файл как .docx и повторите."
    exit 0
}

if ($directExtensions -notcontains $ext) {
    Show-Message "Файл '$($file.Name)':`nформат '$ext' не поддерживается конвертацией в MD."
    exit 0
}

$mdDir = Join-Path $file.DirectoryName "0.1. MD"
New-Item -ItemType Directory -Path $mdDir -Force | Out-Null
$mdFile = Join-Path $mdDir ($file.BaseName + ".md")

if (Test-Path -LiteralPath $mdFile) {
    # Уже сконвертирован — ничего не делаем и не шумим лишним окном
    exit 0
}

try {
    & $pythonForMd -m markitdown "$($file.FullName)" -o "$mdFile" 2>$null
} catch {
    # markitdown мог написать что-то в stderr и вызвать здесь ошибку PowerShell
    # даже при успешном результате — реальный итог проверяем ниже по факту
    # появления файла, а не по этому исключению.
}
if (-not (Test-Path -LiteralPath $mdFile)) {
    Show-Message "Не удалось сконвертировать '$($file.Name)' в MD.`nПодробности — если запустить эту же команду вручную в PowerShell:`n$pythonForMd -m markitdown `"$($file.FullName)`" -o `"$mdFile`"" "Error"
    exit 1
}
