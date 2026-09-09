<#
    ConvertToMD-OCR.ps1
    Вызывается из контекстного меню проводника (пункт "Конвертировать в MD (с OCR)").
    Только для .pdf. Сначала прогоняет файл через OCR (ocrmypdf --skip-text),
    затем конвертирует результат в .md через markitdown.

    Результаты (как в папочном скрипте pdf_to_md_md_and_not_ocr.ps1):
      <имя>_ocr.pdf         — остаётся рядом, в исходной папке (это единственный
                              PDF, который в ней остаётся)
      0.1. MD\<имя>.md      — итоговый Markdown
      0.2. Not ocr\<имя>.pdf — исходный PDF без OCR-слоя, перенесён сюда

    Исходник переносится в "0.2. Not ocr" ТОЛЬКО после того, как и OCR, и
    .md готовы успешно — если что-то не удалось, исходник остаётся на месте,
    чтобы можно было повторить попытку.

    На успехе скрипт молчит — это важно при OCR сразу нескольких выделенных
    файлов, чтобы не заваливать экран окнами. Об ошибке сообщает окном.

    ВАЖНО про $ErrorActionPreference: специально НЕ ставится в "Stop".
    ocrmypdf и markitdown при обычной успешной работе часто пишут
    диагностические сообщения в stderr — PowerShell может воспринять это
    как завершающую ошибку и оборвать скрипт ДО того, как он дойдёт до
    создания .md, даже если сама конвертация прошла нормально. Поэтому
    вызовы обёрнуты в try/catch, а об успехе/неудаче судим не по коду
    возврата, а по тому, появился ли реально ожидаемый файл.
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Windows.Forms | Out-Null

function Show-Message([string]$text, [string]$icon = "Warning") {
    [System.Windows.Forms.MessageBox]::Show($text, "Конвертация в MD (с OCR)", "OK", $icon) | Out-Null
}

function Test-CommandExists($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    Show-Message "Файл не найден:`n$Path"
    exit 1
}

$file = Get-Item -LiteralPath $Path
$ext  = $file.Extension.ToLower()

if ($ext -ne ".pdf") {
    Show-Message "Файл '$($file.Name)':`nOCR применим только к .pdf. Для этого файла используйте пункт 'Конвертировать в MD' (без OCR)."
    exit 0
}

if ($file.Name -like "*_ocr.pdf") {
    Show-Message "Файл '$($file.Name)':`nпохоже, это уже результат OCR (оканчивается на _ocr.pdf). Запустите OCR на исходном файле."
    exit 0
}

if (Test-CommandExists "python3") {
    $pythonForOcr = "python3"
} elseif (Test-CommandExists "python") {
    $pythonForOcr = "python"
} else {
    Show-Message "Не найдена команда 'python3' (и 'python') в PATH.`nПроверьте установку Python / ocrmypdf." "Error"
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

$env:PYTHONWARNINGS = "ignore::RuntimeWarning:pydub.utils"

$mdDir     = Join-Path $file.DirectoryName "0.1. MD"
$notOcrDir = Join-Path $file.DirectoryName "0.2. Not ocr"
New-Item -ItemType Directory -Path $mdDir     -Force | Out-Null
New-Item -ItemType Directory -Path $notOcrDir -Force | Out-Null

$mdFile     = Join-Path $mdDir ($file.BaseName + ".md")
$ocrFile    = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.pdf")
$notOcrFile = Join-Path $notOcrDir $file.Name

$mdAlreadyDone = Test-Path -LiteralPath $mdFile

if (-not $mdAlreadyDone -and -not (Test-Path -LiteralPath $ocrFile)) {
    try {
        & $pythonForOcr -m ocrmypdf --skip-text -l rus+eng "$($file.FullName)" "$ocrFile" 2>$null
    } catch {
        # ocrmypdf мог написать что-то в stderr и вызвать здесь ошибку PowerShell
        # даже при успешном результате — реальный итог проверяем ниже по факту
        # появления файла, а не по этому исключению.
    }
    if (-not (Test-Path -LiteralPath $ocrFile)) {
        Show-Message "Не удалось выполнить OCR для '$($file.Name)'.`nПодробности — если запустить эту же команду вручную в PowerShell:`n$pythonForOcr -m ocrmypdf --skip-text -l rus+eng `"$($file.FullName)`" `"$ocrFile`"" "Error"
        exit 1
    }
}

if (-not $mdAlreadyDone) {
    try {
        & $pythonForMd -m markitdown "$ocrFile" -o "$mdFile" 2>$null
    } catch {
        # см. комментарий выше — итог проверяем по факту появления .md
    }
    if (-not (Test-Path -LiteralPath $mdFile)) {
        Show-Message "OCR прошёл, но конвертация в MD не удалась для '$($file.Name)'.`nПодробности — если запустить эту же команду вручную в PowerShell:`n$pythonForMd -m markitdown `"$ocrFile`" -o `"$mdFile`"" "Error"
        exit 1
    }
}

# Перенос исходника в "0.2. Not ocr" — только теперь, когда и OCR, и .md точно готовы.
# Если файл уже перенесён (повторный клик) или уже отсутствует — тихо пропускаем.
if ((Test-Path -LiteralPath $file.FullName) -and -not (Test-Path -LiteralPath $notOcrFile)) {
    try {
        Move-Item -LiteralPath $file.FullName -Destination $notOcrFile -ErrorAction Stop
    } catch {
        Show-Message "OCR и MD готовы, но не удалось перенести исходник '$($file.Name)' в '0.2. Not ocr'.`nПричина: $($_.Exception.Message)" "Error"
        exit 1
    }
}
