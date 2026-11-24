param(
    [Parameter(Mandatory = $true)][string]$JobsPath,
    [Parameter(Mandatory = $true)][string]$OutPath,
    [int]$Width = 1700
)
$ErrorActionPreference = 'Stop'

# Windows built-in OCR (offline). Input/output are UTF-8 JSON.
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.Streams.IRandomAccessStream, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
[Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$methods = [System.WindowsRuntimeSystemExtensions].GetMethods()
$asTaskOp = ($methods | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
$asTaskAction = ($methods | Where-Object {
    $_.Name -eq 'AsTask' -and (-not $_.IsGenericMethod) -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
})[0]

function Await($op, $type) {
    $task = $asTaskOp.MakeGenericMethod($type).Invoke($null, @($op))
    $task.Wait(-1) | Out-Null
    $task.Result
}
function AwaitAction($action) {
    $task = $asTaskAction.Invoke($null, @($action))
    $task.Wait(-1) | Out-Null
}

$lang = New-Object Windows.Globalization.Language 'zh-Hans-CN'
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $engine) { throw 'No OCR engine available' }

$tmpRoot = Join-Path $env:TEMP ('audit_ocr_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpRoot | Out-Null

function Ocr-BitmapFile([string]$imagePath) {
    $sf = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imagePath)) ([Windows.Storage.StorageFile])
    $stream = Await ($sf.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bmp = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $result = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
        return $result.Text
    } finally {
        $stream.Dispose()
    }
}

$jobsText = [System.IO.File]::ReadAllText($JobsPath, [System.Text.Encoding]::UTF8)
$jobs = $jobsText | ConvertFrom-Json
$allResults = @()

foreach ($job in $jobs) {
    $item = [ordered]@{ id = [string]$job.id; pages = @(); error = '' }
    try {
        if ($job.kind -eq 'pdf') {
            $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync([string]$job.path)) ([Windows.Storage.StorageFile])
            $pdf = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
            foreach ($pageNo in $job.pages) {
                $png = Join-Path $tmpRoot (([Guid]::NewGuid().ToString('N')) + '.png')
                try {
                    $page = $pdf.GetPage([int]$pageNo - 1)
                    $stream = [System.IO.File]::Create($png)
                    $outStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($stream)
                    try {
                        $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
                        $opts.DestinationWidth = $Width
                        AwaitAction ($page.RenderToStreamAsync($outStream, $opts))
                    } finally {
                        $outStream.Dispose()
                        $stream.Dispose()
                    }
                    $text = Ocr-BitmapFile $png
                    $item.pages += [ordered]@{ no = [int]$pageNo; text = [string]$text }
                } catch {
                    $item.pages += [ordered]@{ no = [int]$pageNo; text = ''; error = [string]$_.Exception.Message }
                } finally {
                    if (Test-Path -LiteralPath $png) { Remove-Item -LiteralPath $png -Force -ErrorAction SilentlyContinue }
                }
            }
        } else {
            $text = Ocr-BitmapFile ([string]$job.path)
            $item.pages += [ordered]@{ no = 1; text = [string]$text }
        }
    } catch {
        $item.error = [string]$_.Exception.Message
    }
    $allResults += $item
}

$json = [ordered]@{ ok = $true; engine = $engine.RecognizerLanguage.LanguageTag; results = $allResults } | ConvertTo-Json -Depth 8 -Compress
[System.IO.File]::WriteAllText($OutPath, $json, (New-Object System.Text.UTF8Encoding($false)))
Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
