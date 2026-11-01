[CmdletBinding()]
param([switch]$Test, [switch]$Package, [switch]$PackageOnly, [string]$OutputDirectory='bin\Release')
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$out = if([IO.Path]::IsPathRooted($OutputDirectory)){[IO.Path]::GetFullPath($OutputDirectory)}else{[IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))}
if ($PackageOnly -and $Test) { throw '-PackageOnly 仅打包已验收构建，不能同时指定 -Test。' }
if ($PackageOnly) {
    $Package=$true
    if (!(Test-Path -LiteralPath (Join-Path $out 'OfficeImageViewer.dll'))) { throw '未找到已有构建，请先运行 build.ps1 -Test。' }
}
if (!$PackageOnly) {
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
if (!(Test-Path (Join-Path $framework 'csc.exe'))) { $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319' }
$csc = Join-Path $framework 'csc.exe'
if (!(Test-Path $csc)) { throw '需要 Windows .NET Framework 4.8 编译器。' }
New-Item -ItemType Directory -Path $out -Force | Out-Null
$refs = @('System.dll','System.Xaml.dll','System.Core.dll','System.Xml.dll','System.Xml.Linq.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','Microsoft.CSharp.dll','System.Drawing.dll')
$wpf = @('WindowsBase.dll','PresentationCore.dll','PresentationFramework.dll','UIAutomationTypes.dll','UIAutomationProvider.dll')
$common = @('/nologo','/utf8output','/langversion:5','/optimize+','/warn:4','/platform:anycpu')
foreach ($r in $refs) { $common += '/reference:' + (Join-Path $framework $r) }
foreach ($r in $wpf) { $common += '/reference:' + (Join-Path (Join-Path $framework 'WPF') $r) }
$sources = @(Get-ChildItem (Join-Path $root 'src') -Filter '*.cs' -Recurse | ForEach-Object FullName)
& $csc @common /target:library ("/out:" + (Join-Path $out 'OfficeImageViewer.dll')) @sources
if ($LASTEXITCODE -ne 0) { throw '插件编译失败。' }
$tests = Join-Path $root 'tests\CoreTests.cs'
if (Test-Path $tests) {
    $core = @(Get-ChildItem (Join-Path $root 'src\Core') -Filter '*.cs' | ForEach-Object FullName)
    & $csc @common /target:exe ("/out:" + (Join-Path $out 'CoreTests.exe')) @core $tests
    if ($LASTEXITCODE -ne 0) { throw '核心测试编译失败。' }
}
$hostTests = Join-Path $root 'tests\HostAdapterTests.cs'
if (Test-Path $hostTests) {
    & $csc @common /target:exe /main:HostAdapterTests ("/out:" + (Join-Path $out 'HostAdapterTests.exe')) ("/reference:" + (Join-Path $out 'OfficeImageViewer.dll')) $hostTests $tests
    if ($LASTEXITCODE -ne 0) { throw '宿主适配测试编译失败。' }
}
$harness = Join-Path $root 'tests\ViewerHarness.cs'
if (Test-Path $harness) {
    & $csc @common /target:exe ("/out:" + (Join-Path $out 'ViewerHarness.exe')) ("/reference:" + (Join-Path $out 'OfficeImageViewer.dll')) $harness
    if ($LASTEXITCODE -ne 0) { throw '查看器验证程序编译失败。' }
}
Write-Host "构建成功：$out"
    $setupSources = @(Get-ChildItem (Join-Path $root 'setup') -Filter '*.cs' | ForEach-Object FullName)
    if ($setupSources.Count) {
        & $csc @common /target:winexe ("/out:" + (Join-Path $out 'OfficeImageViewer.Setup.exe')) ("/win32manifest:" + (Join-Path $root 'setup\app.manifest')) ("/resource:" + (Join-Path $out 'OfficeImageViewer.dll') + ',OfficeImageViewer.dll') ("/resource:" + (Join-Path $root 'LICENSE') + ',LICENSE') @setupSources
        if ($LASTEXITCODE -ne 0) { throw '安装程序编译失败。' }
        & $csc @common /target:exe ("/out:" + (Join-Path $out 'SetupEngineTests.exe')) (Join-Path $root 'setup\RegistryStore.cs') (Join-Path $root 'setup\SetupEngine.cs') (Join-Path $root 'tests\SetupEngineTests.cs')
        if ($LASTEXITCODE -ne 0) { throw '安装引擎测试编译失败。' }
        & $csc @common /target:winexe ("/out:" + (Join-Path $out 'SetupUiHarness.exe')) ("/reference:" + (Join-Path $out 'OfficeImageViewer.Setup.exe')) (Join-Path $root 'tests\SetupUiHarness.cs')
        if ($LASTEXITCODE -ne 0) { throw '安装界面验证程序编译失败。' }
    }
if ($Test) {
    & (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'tests\ScriptCompatibility.ps1')
    if ($LASTEXITCODE -ne 0) { throw 'Windows PowerShell 5.1 script compatibility check failed.' }
    & (Join-Path $out 'CoreTests.exe')
    if ($LASTEXITCODE -ne 0) { throw '核心测试失败。' }
    & (Join-Path $out 'HostAdapterTests.exe')
    if ($LASTEXITCODE -ne 0) { throw '宿主适配测试失败。' }
    if ($setupSources.Count) {
        & (Join-Path $out 'SetupEngineTests.exe') (Join-Path $out 'OfficeImageViewer.dll')
        if ($LASTEXITCODE -ne 0) { throw '安装引擎测试失败。' }
    }
}
}
if ($Package) {
    $dist = Join-Path $root 'dist'
    $stage = Join-Path $dist 'OfficeImageViewer'
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    Copy-Item (Join-Path $out 'OfficeImageViewer.dll') $stage -Force
    foreach ($name in @('install.ps1','uninstall.ps1','diagnose.ps1','verify-office.ps1','README.md','LICENSE')) { Copy-Item (Join-Path $root $name) $stage -Force }
    $releaseVersion=[Reflection.AssemblyName]::GetAssemblyName((Join-Path $out 'OfficeImageViewer.dll')).Version.ToString(3)
    $releaseZip=Join-Path $dist ("OfficeImageViewer-$releaseVersion.zip")
    Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $releaseZip -Force
    Write-Host "发布包：$releaseZip"
    $setupExe=Join-Path $out 'OfficeImageViewer.Setup.exe'
    if(Test-Path -LiteralPath $setupExe){Copy-Item -LiteralPath $setupExe -Destination (Join-Path $dist "OfficeImageViewer-Setup-$releaseVersion.exe") -Force}
}




