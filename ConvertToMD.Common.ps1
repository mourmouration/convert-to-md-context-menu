<#
    ConvertToMD.Common.ps1
    Общие функции для ConvertToMD.ps1 и ConvertToMD-OCR.ps1.

    Этот файл сам по себе ничего не делает — он подключается (dot-source)
    из обоих рабочих скриптов и должен лежать с ними в одной папке.

    Здесь собрано всё, от чего зависит корректность работы:
      - запуск внешней программы с проверкой кода завершения и сохранением
        диагностики (а не только «появился ли файл»);
      - расшифровка кодов завершения ocrmypdf на человеческий язык;
      - выбор имени .md с учётом конфликта (Договор.pdf и Договор.docx
        в одной папке больше не претендуют на один и тот же файл);
      - проверка, не устарел ли уже готовый результат;
      - запись результата через временный файл, чтобы после сбоя не
        оставался обрубок, который в следующий раз примут за готовый.
#>

Add-Type -AssemblyName System.Windows.Forms | Out-Null

# Форматы, которые markitdown конвертирует напрямую.
# Внимание: .doc и .ppt (форматы Office 97-2003) markitdown не поддерживает —
# для них в markitdown нет конвертера, они вынесены в отдельный список.
$script:DirectExtensions      = @(".pdf", ".docx", ".xlsx", ".xls", ".pptx", ".html", ".htm", ".csv", ".txt", ".epub")
$script:UnsupportedExtensions = @(".doc", ".ppt")

$script:MdFolderName     = "0.1. MD"
$script:NotOcrFolderName = "0.2. Not ocr"

# Заголовок окон с сообщениями — каждый скрипт выставляет свой
$script:DialogTitle = "Конвертация в MD"

function Show-Message {
    param(
        [string]$Text,
        [string]$Icon = "Warning"
    )
    [System.Windows.Forms.MessageBox]::Show($Text, $script:DialogTitle, "OK", $Icon) | Out-Null
}

function Test-CommandExists {
    param([string]$Name)
    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

<#
    Выбор команды Python. На разных компьютерах Python доступен то как
    'python', то как 'python3', иногда как обе сразу — и тогда они могут
    указывать на разные установки с разным набором пакетов. Поэтому мало
    проверить наличие команды: нужно убедиться, что в этом интерпретаторе
    действительно есть нужный модуль. Проверяем импортом.
#>
function Resolve-PythonForModule {
    param(
        [string]$Module,
        [string[]]$Candidates
    )
    # Сначала ищем интерпретатор, в котором модуль реально импортируется
    foreach ($candidate in $Candidates) {
        if (-not (Test-CommandExists $candidate)) { continue }
        try {
            $null = & $candidate -c "import $Module" 2>&1
            if ($LASTEXITCODE -eq 0) { return $candidate }
        } catch {
            # интерпретатор есть, но запустить не удалось — пробуем следующий
        }
    }
    # Модуль не найден нигде — возвращаем первый существующий интерпретатор,
    # чтобы вызывающий код смог показать осмысленную ошибку с деталями
    foreach ($candidate in $Candidates) {
        if (Test-CommandExists $candidate) { return $candidate }
    }
    return $null
}

<#
    Запуск внешней программы.

    Диагностику (stderr) НЕ выбрасываем, а сохраняем: именно её отсутствие
    раньше превращало любой сбой в бесполезное «не удалось, запустите
    вручную». Код завершения тоже проверяем — у ocrmypdf есть коды, при
    которых выходной файл создан, но некорректен, так что судить об успехе
    по одному лишь наличию файла нельзя.

    $ErrorActionPreference намеренно оставлен в "Continue" в вызывающих
    скриптах: ocrmypdf и markitdown пишут в stderr и при нормальной работе,
    и при "Stop" PowerShell оборвал бы скрипт на ровном месте.
#>
function Invoke-ExternalTool {
    param(
        [string]$Exe,
        [string[]]$ToolArgs,
        [int[]]$OkExitCodes = @(0)
    )

    $captured = @()
    $code = -1
    try {
        $captured = & $Exe @ToolArgs 2>&1
        $code = $LASTEXITCODE
    } catch {
        $captured = @($_.Exception.Message)
        $code = -1
    }

    $lines = @()
    foreach ($item in $captured) {
        if ($null -ne $item) {
            $text = $item.ToString()
            if ($text.Trim() -ne "") { $lines += $text }
        }
    }
    # Для окна с ошибкой хватает хвоста — там обычно и лежит причина
    $tail = @($lines | Select-Object -Last 12)

    return [PSCustomObject]@{
        ExitCode = $code
        Ok       = ($OkExitCodes -contains $code)
        Details  = ($tail -join "`r`n")
    }
}

# Расшифровка кодов завершения ocrmypdf (документированный список)
function Get-OcrExitCodeText {
    param([int]$Code)
    switch ($Code) {
        1   { return "неверные параметры запуска" }
        2   { return "файл не является корректным PDF" }
        3   { return "не установлен требуемый внешний компонент (Tesseract или Ghostscript)" }
        4   { return "результат создан, но не является корректным PDF" }
        5   { return "нет прав доступа к файлу или папке" }
        6   { return "в файле уже есть текстовый слой — OCR не требуется" }
        7   { return "сбой внешней программы (Tesseract или Ghostscript)" }
        8   { return "PDF зашифрован или защищён паролем" }
        9   { return "Tesseract отклонил параметры — проверьте языковые пакеты rus и eng" }
        15  { return "внутренняя ошибка ocrmypdf" }
        130 { return "обработка прервана" }
        -1  { return "не удалось запустить программу" }
        default { return "код завершения $Code" }
    }
}

<#
    Имя итогового .md.

    Раньше имя строилось только из имени без расширения, поэтому Договор.pdf
    и Договор.docx в одной папке претендовали на один и тот же Договор.md —
    второй документ молча терялся. Теперь: если рядом лежит другой документ
    с тем же именем, но иным расширением, к имени добавляется расширение
    исходника (Договор.pdf.md и Договор.docx.md). Если конфликта нет, имя
    остаётся прежним — Договор.md, — чтобы уже сконвертированные архивы не
    пришлось переделывать.
#>
function Get-MdTargetPath {
    param(
        [System.IO.FileInfo]$SourceFile,
        [string]$MdDir
    )

    $knownExtensions = $script:DirectExtensions + $script:UnsupportedExtensions

    $competitors = @(
        Get-ChildItem -LiteralPath $SourceFile.DirectoryName -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.BaseName -eq $SourceFile.BaseName -and
                $_.Name -ne $SourceFile.Name -and
                ($knownExtensions -contains $_.Extension.ToLower())
            }
    )

    $hasConflict = ($competitors.Count -gt 0)

    # Отдельный случай: PDF мог быть уже обработан и перенесён в
    # "0.2. Not ocr", и тогда рядом его больше нет — остался только
    # <имя>_ocr.pdf. Он тоже означает, что документ с таким именем есть,
    # иначе Договор.docx получил бы то Договор.docx.md (пока PDF лежал
    # рядом), то Договор.md (после переноса) — в зависимости от момента
    # запуска. Для самого PDF его собственный _ocr.pdf конфликтом не считаем.
    if (-not $hasConflict -and $SourceFile.Extension.ToLower() -ne ".pdf") {
        $ocrTwin = Join-Path $SourceFile.DirectoryName ($SourceFile.BaseName + "_ocr.pdf")
        if (Test-Path -LiteralPath $ocrTwin) { $hasConflict = $true }
    }

    if ($hasConflict) {
        return (Join-Path $MdDir ($SourceFile.Name + ".md"))
    }
    return (Join-Path $MdDir ($SourceFile.BaseName + ".md"))
}

# Результат считается устаревшим, если его нет или исходник новее его
function Test-OutputIsStale {
    param(
        [string]$OutputPath,
        [string]$InputPath
    )
    if (-not (Test-Path -LiteralPath $OutputPath)) { return $true }
    if (-not (Test-Path -LiteralPath $InputPath))  { return $false }

    $outputItem = Get-Item -LiteralPath $OutputPath
    $sourceItem = Get-Item -LiteralPath $InputPath
    return ($sourceItem.LastWriteTimeUtc -gt $outputItem.LastWriteTimeUtc)
}

<#
    Путь для временного файла рядом с итоговым. Пишем сначала сюда, и только
    после успешной проверки переименовываем в итоговый файл. Так после сбоя
    рядом не остаётся обрубка, который в следующий раз сочтут готовым
    результатом. Файл создаётся в той же папке, поэтому переименование
    выполняется мгновенно, без копирования между дисками.
#>
function New-TempOutputPath {
    param([string]$FinalPath)

    $dir       = Split-Path -Parent $FinalPath
    $leaf      = Split-Path -Leaf $FinalPath
    $extension = [System.IO.Path]::GetExtension($leaf)
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    $stamp     = "{0}-{1}" -f $PID, (Get-Random -Maximum 1000000)

    return (Join-Path $dir ("~$baseName.partial-$stamp$extension"))
}

# Удаление собственного временного файла (только его — чужие файлы не трогаем)
function Remove-TempOutput {
    param([string]$Path)
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

<#
    Свободное имя для переноса. Если в папке назначения уже лежит файл с
    таким именем, подбираем «Имя (2).pdf», «Имя (3).pdf» и так далее.
    Существующий файл никогда не перезаписывается.
#>
function Get-NonCollidingPath {
    param([string]$DesiredPath)

    if (-not (Test-Path -LiteralPath $DesiredPath)) { return $DesiredPath }

    $dir       = Split-Path -Parent $DesiredPath
    $leaf      = Split-Path -Leaf $DesiredPath
    $extension = [System.IO.Path]::GetExtension($leaf)
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($leaf)

    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $dir ("$baseName ($i)$extension")
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    return $null
}

# Проверка, что созданный .md не пустой. Пустой результат означает, что текст
# из документа извлечь не удалось — молча выдавать такой файл нельзя.
function Test-MarkdownHasContent {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }

    $file = Get-Item -LiteralPath $Path
    if ($file.Length -eq 0) { return $false }
    if ($file.Length -gt 512) { return $true }

    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
    if ($null -eq $content) { return $false }
    return ($content.Trim().Length -gt 0)
}
