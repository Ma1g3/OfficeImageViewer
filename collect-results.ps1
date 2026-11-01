[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core'){throw '请使用 Windows PowerShell 5.1：powershell -NoProfile -File .\collect-results.ps1'}
if(!$OutputDirectory){$OutputDirectory=Join-Path $PSScriptRoot 'validation-results'}
$reportName='OfficeImageViewer-results-'+[DateTime]::Now.ToString('yyyyMMdd-HHmmss')+'-'+[Guid]::NewGuid().ToString('N').Substring(0,8)
$reportDir=Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) $reportName
New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
$environment=Join-Path $reportDir 'office-environment.json'
& (Join-Path $PSScriptRoot 'verify-office.ps1') -OutputPath $environment -OmitUser
if(!(Test-Path -LiteralPath $environment)){throw '未能生成 Office 环境快照。'}
$system=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$office=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
$additional=[ordered]@{
    WindowsProduct=$system.ProductName
    WindowsDisplayVersion=$system.DisplayVersion
    WindowsBuild=$system.CurrentBuild
    WindowsUBR=$system.UBR
    Is64BitOperatingSystem=[Environment]::Is64BitOperatingSystem
    CollectorProcessBits=[IntPtr]::Size*8
    OfficePlatform=$office.Platform
    OfficeFullVersion=$office.VersionToReport
    ValidationResult='Manual checks required; collection does not prove host behavior.'
}
$additional | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $reportDir 'system.json') -Encoding UTF8
$checklist=Join-Path $PSScriptRoot 'validation-checklist.md'
if(Test-Path -LiteralPath $checklist){Copy-Item -LiteralPath $checklist -Destination $reportDir}
$zip=$reportDir+'.zip'
Compress-Archive -LiteralPath @(Get-ChildItem -LiteralPath $reportDir -File | ForEach-Object FullName) -DestinationPath $zip
Write-Output "结果包已生成：$zip"
Write-Output '仅采集环境元数据及验证清单，没有收集文档、图片或截图，也没有自动上传。'
