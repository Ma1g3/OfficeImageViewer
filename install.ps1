[CmdletBinding(SupportsShouldProcess=$true)]
param([ValidateSet('Both','Word','PowerPoint')][string]$HostApp='Both', [switch]$AllowMissingOffice, [ValidateSet('CurrentUser','AllUsers')][string]$Scope='CurrentUser')
$ErrorActionPreference='Stop'
$hive=if($Scope -eq 'AllUsers'){[Microsoft.Win32.RegistryHive]::LocalMachine}else{[Microsoft.Win32.RegistryHive]::CurrentUser}
$scopeLabel=if($Scope -eq 'AllUsers'){'所有用户'}else{'当前用户'}
if($Scope -eq 'AllUsers' -and !$WhatIfPreference -and ![Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'AllUsers 安装/卸载需要管理员权限 PowerShell。'}
$clsid='{26C7344A-8129-4C5B-A462-034F9B5FC712}'
$progId='OfficeImageViewer.AddIn'
$dll=Join-Path $PSScriptRoot 'OfficeImageViewer.dll'
if (!(Test-Path -LiteralPath $dll)) { $dll=Join-Path $PSScriptRoot 'bin\Release\OfficeImageViewer.dll' }
if (!(Test-Path -LiteralPath $dll)) { throw '未找到插件 DLL，请先运行 build.ps1 或解压完整发布包。' }
$assemblyName=[Reflection.AssemblyName]::GetAssemblyName($dll)
$identity=$assemblyName.FullName
$codeBase=$assemblyName.CodeBase
$apps=if($HostApp -eq 'Both'){@('Word','PowerPoint')}else{@($HostApp)}
$processes=@('WINWORD','POWERPNT')
if (Get-Process -Name $processes -ErrorAction SilentlyContinue) { throw '请先关闭 Word 和 PowerPoint，再运行安装。' }
$release=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
if (!$release -or $release -lt 528040) { throw '需要 .NET Framework 4.8 或更高版本。' }
foreach($app in $apps) {
    if (!(Test-Path "Registry::HKEY_CLASSES_ROOT\$app.Application") -and !$AllowMissingOffice) {
        throw "未检测到 $app 桌面版。请安装 Office 后重试；仅测试注册时可使用 -AllowMissingOffice。"
    }
}
if (!$PSCmdlet.ShouldProcess("$progId ($scopeLabel, $($apps -join ', '))",'注册插件')) { return }
$views=@([Microsoft.Win32.RegistryView]::Registry32)
if([Environment]::Is64BitOperatingSystem){$views += [Microsoft.Win32.RegistryView]::Registry64}
# Preflight every registry view before changing any of them.
foreach($view in $views) {
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try {
        $existing=$base.OpenSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
        if($existing){
            try {$old=$existing.GetValue('CodeBase')}finally{$existing.Dispose()}
            if($old -and $old -ne $codeBase){throw "已有其他位置的插件注册：$old。请先从旧位置卸载。"}
        }
    } finally {$base.Dispose()}
}
foreach($view in $views) {
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try {

        $entries=@{
            "Software\Classes\$progId"=@{''='Office Image Viewer'};
            "Software\Classes\$progId\CLSID"=@{''=$clsid};
            "Software\Classes\CLSID\$clsid"=@{''='OfficeImageViewer.AddIn'};
            "Software\Classes\CLSID\$clsid\ProgId"=@{''=$progId};
            "Software\Classes\CLSID\$clsid\InprocServer32"=@{''='mscoree.dll';'ThreadingModel'='Both';'Class'='OfficeImageViewer.AddIn';'Assembly'=$identity;'RuntimeVersion'='v4.0.30319';'CodeBase'=$codeBase}
        }
        foreach($path in $entries.Keys){$key=$base.CreateSubKey($path);try{foreach($n in $entries[$path].Keys){$key.SetValue($n,$entries[$path][$n],[Microsoft.Win32.RegistryValueKind]::String)}}finally{$key.Dispose()}}
        $version=[Reflection.AssemblyName]::GetAssemblyName($dll).Version.ToString()
        $key=$base.CreateSubKey("Software\Classes\CLSID\$clsid\InprocServer32\$version")
        try{foreach($n in @('Class','Assembly','RuntimeVersion','CodeBase')){$key.SetValue($n,$entries["Software\Classes\CLSID\$clsid\InprocServer32"][$n],[Microsoft.Win32.RegistryValueKind]::String)}}finally{$key.Dispose()}
        $key=$base.CreateSubKey("Software\Classes\CLSID\$clsid\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}");$key.Dispose()
        foreach($app in $apps){
            $key=$base.CreateSubKey("Software\Microsoft\Office\$app\Addins\$progId")
            try {
                $key.SetValue('FriendlyName','Office 图片查看器')
                $key.SetValue('Description','双击 Word / PowerPoint 图片放大查看')
                $key.SetValue('LoadBehavior',3,[Microsoft.Win32.RegistryValueKind]::DWord)
            } finally {$key.Dispose()}
        }
    } finally {$base.Dispose()}
}
Write-Host '安装完成。请保留当前插件目录，重新打开 Word / PowerPoint 后查看“图片查看”选项卡。'



