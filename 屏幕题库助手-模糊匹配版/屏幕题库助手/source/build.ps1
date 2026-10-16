$ErrorActionPreference = 'Stop'
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
$output = Split-Path $PSScriptRoot -Parent
& (Join-Path $framework 'csc.exe') /nologo /target:winexe /optimize+ /codepage:65001 /win32manifest:"$PSScriptRoot\app.manifest" /utf8output /out:"$output\屏幕题库助手.exe" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Xml.Linq.dll /reference:System.IO.Compression.dll /reference:System.IO.Compression.FileSystem.dll /reference:System.Web.Extensions.dll "$PSScriptRoot\Core.cs" "$PSScriptRoot\App.cs" "$PSScriptRoot\SelfTest.cs"
if ($LASTEXITCODE -ne 0) { throw '编译失败' }
