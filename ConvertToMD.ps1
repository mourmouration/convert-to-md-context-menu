<#
    ConvertToMD.ps1
    Вызывается из контекстного меню проводника (пункт "Конвертировать в MD").
    Конвертирует ОДИН файл напрямую через markitdown, без OCR.
    Результат кладётся в подпапку "0.1. MD" рядом с исходным файлом.

    Исходный файл не перемещается и не изменяется.

    Поддерживаются: .pdf, .docx, .xlsx, .xls, .pptx, .html, .htm, .csv,
    .txt, .epub. Для .pdf без текстового слоя (сканы) результат окажется
    пустым — на это скрипт отдельно указывает и предлагает пункт
    "Конвертировать в MD (с OCR)".
    Форматы Office 97-2003 (.doc, .ppt) markitdown не поддерживает —
    в нём попросту нет конвертера для этих форматов.

    Повторный запуск: если .md уже есть и исходник с тех пор не менялся,
    скрипт молча выходит. Если исходник изменился — .md пересоздаётся.
    Раньше существующий .md считался готовым результатом безусловно, и
    после правки документа конвертация молча возвращала устаревший текст.

    На успехе скрипт молчит — это важно при конвертации сразу нескольких
    выделенных файлов, чтобы не заваливать экран окнами. Об ошибке или
    неподдерживаемом формате сообщает окном, с указанием причины.
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

# Намеренно НЕ "Stop": markitdown пишет в stderr и при нормальной работе,
# а PowerShell при "Stop" воспринял бы это как завершающую ошибку и оборвал
# скрипт до создания результата. Об успехе судим по коду завершения и по
# фактической проверке результата — см. ConvertToMD.Common.ps1.
$ErrorActionPreference = "Continue"

$commonScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "ConvertToMD.Common.ps1"
if (-not (Test-Path -LiteralPath $commonScript)) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show(
        "Не найден файл ConvertToMD.Common.ps1.`nОн должен лежать в одной папке с этим скриптом.",
        "Конвертация в MD", "OK", "Error") | Out-Null
    exit 1
}
. $commonScript

$script:DialogTitle = "Конвертация в MD"

# --- Проверка входного файла ---

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    Show-Message "Файл не найден:`r`n$Path"
    exit 1
}

$file = Get-Item -LiteralPath $Path
$extension = $file.Extension.ToLower()

if ($script:UnsupportedExtensions -contains $extension) {
    $modern = if ($extension -eq ".doc") { ".docx" } else { ".pptx" }
    Show-Message "Файл '$($file.Name)':`r`nформат '$extension' (Office 97-2003) markitdown не поддерживает.`r`n`r`nПересохраните файл как $modern и повторите."
    exit 0
}

if ($script:DirectExtensions -notcontains $extension) {
    Show-Message "Файл '$($file.Name)':`r`nформат '$extension' не поддерживается конвертацией в MD."
    exit 0
}

# --- Выбор интерпретатора ---
# Проверяем не наличие команды, а наличие в ней модуля markitdown: на
# компьютере может стоять несколько установок Python.

$pythonForMd = Resolve-PythonForModule -Module "markitdown" -Candidates @("python", "python3")
if (-not $pythonForMd) {
    Show-Message "Не найден Python (ни 'python', ни 'python3') в PATH.`r`nУстановите Python и пакет markitdown." "Error"
    exit 1
}

# Заглушаем безобидное предупреждение pydub про отсутствующий ffmpeg —
# markitdown подгружает этот модуль всегда, хотя для документов он не нужен.
$env:PYTHONWARNINGS = "ignore::RuntimeWarning:pydub.utils"

# --- Пути ---

$mdDir  = Join-Path $file.DirectoryName $script:MdFolderName
$mdFile = Get-MdTargetPath -SourceFile $file -MdDir $mdDir

# Уже сконвертирован и исходник с тех пор не менялся — тихо выходим
if (-not (Test-OutputIsStale -OutputPath $mdFile -InputPath $file.FullName)) {
    exit 0
}

# --- Конвертация ---

try {
    New-Item -ItemType Directory -Path $mdDir -Force -ErrorAction Stop | Out-Null
} catch {
    Show-Message "Не удалось создать папку '$($script:MdFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
    exit 1
}

# Пишем во временный файл рядом и переименовываем только после проверки —
# чтобы после сбоя не остался обрубок, который в следующий раз сочтут готовым
$tempMd = New-TempOutputPath -FinalPath $mdFile
Remove-TempOutput -Path $tempMd

$result = Invoke-ExternalTool `
    -Exe $pythonForMd `
    -ToolArgs @("-m", "markitdown", $file.FullName, "-o", $tempMd) `
    -OkExitCodes @(0)

if (-not $result.Ok -or -not (Test-Path -LiteralPath $tempMd)) {
    Remove-TempOutput -Path $tempMd
    $message = "Не удалось сконвертировать '$($file.Name)' в MD."
    if ($result.Details) {
        $message += "`r`n`r`nСообщение программы:`r`n$($result.Details)"
    } else {
        $message += "`r`n`r`nПроверьте, что установлены нужные компоненты markitdown:`r`n" +
                    "python -m pip install --upgrade `"markitdown[all]`""
    }
    Show-Message $message "Error"
    exit 1
}

$hasContent = Test-MarkdownHasContent -Path $tempMd

try {
    Move-Item -LiteralPath $tempMd -Destination $mdFile -Force -ErrorAction Stop
} catch {
    Remove-TempOutput -Path $tempMd
    Show-Message "Конвертация прошла, но не удалось сохранить результат в '$($script:MdFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
    exit 1
}

# Пустой результат — не успех. Молча выдавать пустой файл нельзя: по нему
# не видно, что документ на самом деле не прочитан.
if (-not $hasContent) {
    if ($extension -eq ".pdf") {
        Show-Message ("Из файла '$($file.Name)' не извлечено ни одного символа.`r`n`r`n" +
            "Похоже, это скан: текст в нём — изображение. Используйте пункт " +
            "'Конвертировать в MD (с OCR)'.")
    } else {
        Show-Message ("Из файла '$($file.Name)' не извлечено ни одного символа.`r`n`r`n" +
            "Возможно, документ пуст или его содержимое хранится в виде изображений.")
    }
    exit 1
}
