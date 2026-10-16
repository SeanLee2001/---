$ErrorActionPreference = 'Stop'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
function Await($operation, $type) {
    $task = $asTask.MakeGenericMethod($type).Invoke($null, @($operation))
    $task.GetAwaiter().GetResult()
}
try {
    $lang = New-Object Windows.Globalization.Language('zh-Hans-CN')
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    if ($null -eq $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
    if ($null -eq $engine) { throw '未找到 OCR 语言。请在 Windows 设置中安装中文语言的光学字符识别功能。' }
    [Console]::WriteLine((@{ ready=$true; language=$engine.RecognizerLanguage.LanguageTag; maxDimension=[Windows.Media.Ocr.OcrEngine]::MaxImageDimension } | ConvertTo-Json -Compress))
} catch {
    [Console]::WriteLine((@{ error=$_.Exception.Message } | ConvertTo-Json -Compress))
    exit 1
}
while ($null -ne ($line = [Console]::ReadLine())) {
    $stream = $null; $bitmap = $null
    try {
        $watch = [Diagnostics.Stopwatch]::StartNew(); $request = $line | ConvertFrom-Json
        $phase = '读取图片'; $memory = New-Object System.IO.MemoryStream(,[System.IO.File]::ReadAllBytes($request.path))
        $phase = '打开图片'; $stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($memory)
        $phase = '解码图片'; $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $phase = '生成位图'; $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $phase = '文字识别'; $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @($result.Lines | ForEach-Object { $_.Text })
        [Console]::WriteLine((@{ elapsedMs=$watch.ElapsedMilliseconds; text=($lines -join "`n"); language=$engine.RecognizerLanguage.LanguageTag } | ConvertTo-Json -Compress))
    } catch {
        [Console]::WriteLine((@{ error=($phase + ': ' + $_.Exception.Message) } | ConvertTo-Json -Compress))
    } finally {
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}
