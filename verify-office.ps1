[CmdletBinding()]
param([string]$OutputPath, [switch]$OmitUser)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core'){throw '请使用 Windows PowerShell 5.1：powershell -NoProfile -File .\verify-office.ps1'}
if(!$OutputPath){$OutputPath=Join-Path $PSScriptRoot 'work\office-environment.json'}
function Release-Com($obj){if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}}
$results=@()
foreach($name in @('Word','PowerPoint')){
    $hostInfo=[ordered]@{Host=$name;Registered=(Test-Path "Registry::HKEY_CLASSES_ROOT\$name.Application");Running=$false;Version=$null;AddinConnected=$false;OpenDocumentCount=$null;ActiveDocumentSaved=$null;Note=$null}
    $app=$null;$addins=$null;$addin=$null;$documents=$null;$active=$null
    try{
        $app=[Runtime.InteropServices.Marshal]::GetActiveObject("$name.Application")
        $hostInfo.Running=$true
        $hostInfo.Version=[string]$app.Version
        $addins=$app.COMAddIns
        try{$addin=$addins.Item('OfficeImageViewer.AddIn');$hostInfo.AddinConnected=[bool]$addin.Connect}
        catch{$hostInfo.Note='当前 Office 实例未发现已连接的插件。'}
        if($name -eq 'Word'){
            $documents=$app.Documents
            $hostInfo.OpenDocumentCount=$documents.Count
            if($documents.Count -gt 0){$active=$app.ActiveDocument;$hostInfo.ActiveDocumentSaved=[bool]$active.Saved}
        }else{
            $documents=$app.Presentations
            $hostInfo.OpenDocumentCount=$documents.Count
            if($documents.Count -gt 0){$active=$app.ActivePresentation;$hostInfo.ActiveDocumentSaved=[Convert]::ToBoolean($active.Saved)}
        }
    }catch{$hostInfo.Note='未连接到运行中的 Office。请以相同用户和权限打开对应应用后重试。'}
    finally{Release-Com $active;Release-Com $documents;Release-Com $addin;Release-Com $addins;Release-Com $app}
    $results += [pscustomobject]$hostInfo
}
$dll=Join-Path $PSScriptRoot 'OfficeImageViewer.dll'
if(!(Test-Path $dll)){$dll=Join-Path $PSScriptRoot 'bin\Release\OfficeImageViewer.dll'}
$hash=if(Test-Path $dll){(Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash}else{$null}
$report=[ordered]@{
    CapturedUtc=[DateTimeOffset]::UtcNow.ToString('o')
    User=[Security.Principal.WindowsIdentity]::GetCurrent().Name
    Elevated=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    FrameworkRelease=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
    AddinSha256=$hash
    Hosts=$results
    Scope='只读环境和加载状态快照，不代表双击、保存状态不变或兼容性验收通过。'
}
if($OmitUser){$report.Remove('User')}
$full=[IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($full)) -Force | Out-Null
$report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $full -Encoding UTF8
$results | Format-Table Host,Registered,Running,Version,AddinConnected,ActiveDocumentSaved
Write-Output "只读快照已写入：$full"
Write-Output $report.Scope

