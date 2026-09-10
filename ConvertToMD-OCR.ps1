<#
    ConvertToMD-OCR.ps1
    Вызывается из контекстного меню проводника (пункт "Конвертировать в MD (с OCR)").
    Только для .pdf. Сначала прогоняет файл через OCR (ocrmypdf --skip-text),
    затем конвертирует результат в .md через markitdown.

    Результаты:
      <имя>_ocr.pdf          — остаётся в исходной папке (единственный PDF,
                               который в ней остаётся)
      0.1. MD\<имя>.md       — итоговый Markdown
      0.2. Not ocr\<имя>.pdf — исходный PDF без OCR-слоя, перенесён сюда

    ПОРЯДОК ПРОВЕРОК (здесь была исправлена серьёзная ошибка).
    Раньше обе стадии и перенос исходника зависели от одного условия —
    «существует ли итоговый .md». Из-за этого посторонний файл с подходящим
    именем (например, .md, полученный ранее обычной конвертацией того же
    скана) заставлял скрипт пропустить и OCR, и конвертацию, но всё равно
    перенести исходный PDF в "0.2. Not ocr" — в исходной папке не оставалось
    ни одного PDF вообще.

    Теперь каждый шаг проверяется отдельно и по своему входу:
      OCR      выполняется, если _ocr.pdf отсутствует или старше исходника;
      MD       выполняется, если .md отсутствует или старше _ocr.pdf;
      перенос  выполняется, только если оба файла реально существуют и оба
               актуальны. Посторонний .md перенос больше не разрешает.

    Побочный полезный эффект: если раньше по этому скану уже был получен
    пустой .md обычной конвертацией, то после OCR файл _ocr.pdf становится
    новее этого .md — и .md автоматически пересоздаётся уже с текстом.

    На успехе скрипт молчит — это важно при обработке сразу нескольких
    выделенных файлов, чтобы не заваливать экран окнами. Об ошибке сообщает
    окном, с указанием причины.
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path
)

# Намеренно НЕ "Stop": ocrmypdf и markitdown пишут в stderr и при нормальной
# работе, а PowerShell при "Stop" воспринял бы это как завершающую ошибку и
# оборвал скрипт до создания результата. Об успехе судим по коду завершения
# и по фактической проверке результата — см. ConvertToMD.Common.ps1.
$ErrorActionPreference = "Continue"

$commonScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "ConvertToMD.Common.ps1"
if (-not (Test-Path -LiteralPath $commonScript)) {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    [System.Windows.Forms.MessageBox]::Show(
        "Не найден файл ConvertToMD.Common.ps1.`nОн должен лежать в одной папке с этим скриптом.",
        "Конвертация в MD (с OCR)", "OK", "Error") | Out-Null
    exit 1
}
. $commonScript

$script:DialogTitle = "Конвертация в MD (с OCR)"

# --- Проверка входного файла ---

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    Show-Message "Файл не найден:`r`n$Path"
    exit 1
}

$file = Get-Item -LiteralPath $Path
$extension = $file.Extension.ToLower()

if ($extension -ne ".pdf") {
    Show-Message "Файл '$($file.Name)':`r`nOCR применим только к .pdf. Для этого файла используйте пункт 'Конвертировать в MD' (без OCR)."
    exit 0
}

if ($file.Name -like "*_ocr.pdf") {
    Show-Message "Файл '$($file.Name)':`r`nпохоже, это уже результат OCR (имя оканчивается на _ocr.pdf). Запустите OCR на исходном файле."
    exit 0
}

# --- Выбор интерпретаторов ---
# Проверяем не наличие команды, а наличие в ней нужного модуля: на компьютере
# может стоять несколько установок Python, и ocrmypdf с markitdown могут
# оказаться в разных.

$pythonForOcr = Resolve-PythonForModule -Module "ocrmypdf" -Candidates @("python3", "python")
if (-not $pythonForOcr) {
    Show-Message "Не найден Python (ни 'python3', ни 'python') в PATH.`r`nУстановите Python и пакет ocrmypdf." "Error"
    exit 1
}

$pythonForMd = Resolve-PythonForModule -Module "markitdown" -Candidates @("python", "python3")
if (-not $pythonForMd) {
    Show-Message "Не найден Python (ни 'python', ни 'python3') в PATH.`r`nУстановите Python и пакет markitdown." "Error"
    exit 1
}

# Заглушаем безобидное предупреждение pydub про отсутствующий ffmpeg —
# markitdown подгружает этот модуль всегда, хотя для документов он не нужен.
$env:PYTHONWARNINGS = "ignore::RuntimeWarning:pydub.utils"

# --- Пути ---

$mdDir      = Join-Path $file.DirectoryName $script:MdFolderName
$notOcrDir  = Join-Path $file.DirectoryName $script:NotOcrFolderName
$ocrFile    = Join-Path $file.DirectoryName ($file.BaseName + "_ocr.pdf")
$mdFile     = Get-MdTargetPath -SourceFile $file -MdDir $mdDir

# --- Шаг 1: OCR ---
# Условие зависит только от исходника и _ocr.pdf, наличие .md здесь ни при чём.

if (Test-OutputIsStale -OutputPath $ocrFile -InputPath $file.FullName) {

    $tempOcr = New-TempOutputPath -FinalPath $ocrFile
    Remove-TempOutput -Path $tempOcr

    # Код 10 = не удалось привести результат к PDF/A, но обычный PDF получен.
    # Для извлечения текста это полностью пригодный файл, поэтому принимаем.
    $ocrResult = Invoke-ExternalTool `
        -Exe $pythonForOcr `
        -ToolArgs @("-m", "ocrmypdf", "--skip-text", "-l", "rus+eng", $file.FullName, $tempOcr) `
        -OkExitCodes @(0, 10)

    if (-not $ocrResult.Ok -or -not (Test-Path -LiteralPath $tempOcr)) {
        Remove-TempOutput -Path $tempOcr
        $reason = Get-OcrExitCodeText -Code $ocrResult.ExitCode
        $message = "Не удалось выполнить OCR для '$($file.Name)'.`r`n`r`nПричина: $reason."
        if ($ocrResult.Details) {
            $message += "`r`n`r`nСообщение программы:`r`n$($ocrResult.Details)"
        }
        Show-Message $message "Error"
        exit 1
    }

    try {
        Move-Item -LiteralPath $tempOcr -Destination $ocrFile -Force -ErrorAction Stop
    } catch {
        Remove-TempOutput -Path $tempOcr
        Show-Message "OCR выполнен, но не удалось сохранить результат в '$([System.IO.Path]::GetFileName($ocrFile))'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
        exit 1
    }
}

# --- Шаг 2: OCR-PDF -> MD ---
# Условие зависит от _ocr.pdf. Поэтому свежий OCR всегда влечёт пересоздание
# .md, даже если файл с таким именем уже лежал в "0.1. MD".

if (Test-OutputIsStale -OutputPath $mdFile -InputPath $ocrFile) {

    try {
        New-Item -ItemType Directory -Path $mdDir -Force -ErrorAction Stop | Out-Null
    } catch {
        Show-Message "Не удалось создать папку '$($script:MdFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
        exit 1
    }

    $tempMd = New-TempOutputPath -FinalPath $mdFile
    Remove-TempOutput -Path $tempMd

    $mdResult = Invoke-ExternalTool `
        -Exe $pythonForMd `
        -ToolArgs @("-m", "markitdown", $ocrFile, "-o", $tempMd) `
        -OkExitCodes @(0)

    if (-not $mdResult.Ok -or -not (Test-Path -LiteralPath $tempMd)) {
        Remove-TempOutput -Path $tempMd
        $message = "OCR прошёл, но конвертация в MD не удалась для '$($file.Name)'."
        if ($mdResult.Details) {
            $message += "`r`n`r`nСообщение программы:`r`n$($mdResult.Details)"
        }
        Show-Message $message "Error"
        exit 1
    }

    # Пустой результат — не успех: значит, текст распознать не удалось.
    # Сам файл сохраняем (иногда полезно посмотреть), но предупреждаем.
    $hasContent = Test-MarkdownHasContent -Path $tempMd

    try {
        Move-Item -LiteralPath $tempMd -Destination $mdFile -Force -ErrorAction Stop
    } catch {
        Remove-TempOutput -Path $tempMd
        Show-Message "Конвертация прошла, но не удалось сохранить результат в '$($script:MdFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
        exit 1
    }

    if (-not $hasContent) {
        Show-Message ("После OCR из файла '$($file.Name)' не извлечено ни одного символа.`r`n`r`n" +
            "Возможные причины: очень низкое качество скана, текст на другом языке (распознавание идёт по rus+eng), " +
            "либо страницы вставлены как изображения в неподдерживаемом виде.`r`n`r`n" +
            "Исходный файл оставлен на месте.")
        exit 1
    }
}

# --- Шаг 3: перенос исходника в "0.2. Not ocr" ---
# Выполняется, только если оба результата реально существуют и актуальны.
# Это и есть исправление главной ошибки: посторонний .md больше не является
# основанием для переноса документа.

$ocrReady = (Test-Path -LiteralPath $ocrFile) -and -not (Test-OutputIsStale -OutputPath $ocrFile -InputPath $file.FullName)
$mdReady  = (Test-Path -LiteralPath $mdFile)  -and -not (Test-OutputIsStale -OutputPath $mdFile  -InputPath $ocrFile)

if (-not ($ocrReady -and $mdReady)) {
    # Досюда доходим, если оба результата были готовы заранее и ничего не
    # пересчитывалось, но какой-то из них при этом не в порядке. Исходник
    # не трогаем — это безопаснее, чем перенести документ неизвестно почему.
    exit 0
}

# Пустой .md готовым результатом не считается — иначе повторный клик по тому
# же файлу перенёс бы исходник, хотя текст так и не был извлечён.
if (-not (Test-MarkdownHasContent -Path $mdFile)) {
    Show-Message ("Файл '$([System.IO.Path]::GetFileName($mdFile))' в папке '$($script:MdFolderName)' пуст — " +
        "текст из '$($file.Name)' извлечь не удалось.`r`n`r`n" +
        "Исходный файл оставлен на месте.")
    exit 1
}

if (-not (Test-Path -LiteralPath $file.FullName)) {
    # Исходник уже перенесён более ранним запуском
    exit 0
}

try {
    New-Item -ItemType Directory -Path $notOcrDir -Force -ErrorAction Stop | Out-Null
} catch {
    Show-Message "OCR и MD готовы, но не удалось создать папку '$($script:NotOcrFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
    exit 1
}

$notOcrFile = Get-NonCollidingPath -DesiredPath (Join-Path $notOcrDir $file.Name)
if (-not $notOcrFile) {
    Show-Message "OCR и MD готовы, но в папке '$($script:NotOcrFolderName)' не удалось подобрать свободное имя для '$($file.Name)'." "Error"
    exit 1
}

try {
    Move-Item -LiteralPath $file.FullName -Destination $notOcrFile -ErrorAction Stop
} catch {
    Show-Message "OCR и MD готовы, но не удалось перенести исходник '$($file.Name)' в '$($script:NotOcrFolderName)'.`r`n`r`nПричина: $($_.Exception.Message)" "Error"
    exit 1
}
