[CmdletBinding()]
param([switch]$UseExistingBuild)
$ErrorActionPreference='Stop'
$root=$PSScriptRoot
if($UseExistingBuild){
    foreach($name in @('OfficeImageViewer.dll','ViewerHarness.exe','OfficeImageViewer.Setup.exe')){
        if(!(Test-Path -LiteralPath (Join-Path $root ('bin\Release\'+$name)))){throw "缺少已有构建：$name"}
    }
    & (Join-Path $root 'build.ps1') -PackageOnly
}else{& (Join-Path $root 'build.ps1') -Test -Package}
$stage=Join-Path $root ('work\validation-package-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stage -Force | Out-Null
foreach($name in @('OfficeImageViewer.dll','ViewerHarness.exe','OfficeImageViewer.Setup.exe')){
    Copy-Item -LiteralPath (Join-Path $root ('bin\Release\'+$name)) -Destination $stage
}
foreach($name in @('install.ps1','uninstall.ps1','diagnose.ps1','verify-office.ps1','collect-results.ps1','validation-checklist.md','验证指南.md','LICENSE')){
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination $stage
}
Copy-Item -LiteralPath (Join-Path $root 'specs\001-image-preview\office-validation.md') -Destination $stage
$releaseVersion=[Reflection.AssemblyName]::GetAssemblyName((Join-Path $stage 'OfficeImageViewer.dll')).Version.ToString(3)
$zip=Join-Path $root ("dist\OfficeImageViewer-validation-$releaseVersion.zip")
Compress-Archive -LiteralPath @(Get-ChildItem -LiteralPath $stage -File | ForEach-Object FullName) -DestinationPath $zip -Force
Write-Output "跨电脑验证包：$zip"
