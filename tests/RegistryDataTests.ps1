[CmdletBinding()]
param([ValidateSet('CurrentUser','AllUsers')][string]$Scope='CurrentUser')
$ErrorActionPreference='Stop'
$hive=if($Scope -eq 'AllUsers'){[Microsoft.Win32.RegistryHive]::LocalMachine}else{[Microsoft.Win32.RegistryHive]::CurrentUser}
$root=Split-Path $PSScriptRoot -Parent
$clsid='{26C7344A-8129-4C5B-A462-034F9B5FC712}'
$progId='OfficeImageViewer.AddIn'
$views=@([Microsoft.Win32.RegistryView]::Registry32)
if([Environment]::Is64BitOperatingSystem){$views += [Microsoft.Win32.RegistryView]::Registry64}
$roots=@("Software\Classes\CLSID\$clsid","Software\Classes\$progId","Software\Microsoft\Office\Word\Addins\$progId","Software\Microsoft\Office\PowerPoint\Addins\$progId")
function Assert-Clean {
    foreach($view in $views){
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
        try {foreach($path in $roots){$key=$base.OpenSubKey($path);if($key){$key.Dispose();throw "测试要求无现有安装：$view $path"}}}finally{$base.Dispose()}
    }
}
Assert-Clean
& (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope -WhatIf
Assert-Clean
Write-Output 'PASS WhatIf does not register'
if([Environment]::Is64BitOperatingSystem){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,[Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key=$base.CreateSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
        try{$key.SetValue('CodeBase','file:///D:/Other/OfficeImageViewer.dll')}finally{$key.Dispose()}
        $rejected=$false
        try{& (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope}catch{if($_.Exception.Message -like '*已有其他位置*'){$rejected=$true}else{throw}}
        if(!$rejected){throw '未拒绝已有路径冲突'}
        $other=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,[Microsoft.Win32.RegistryView]::Registry32)
        try{
            $key=$other.OpenSubKey("Software\Classes\CLSID\$clsid")
            if($key){$key.Dispose();throw '64 位冲突前错误写入了 32 位注册'}
        }finally{$other.Dispose()}
        $key=$base.OpenSubKey("Software\Classes\CLSID\$clsid\InprocServer32")
        try{if($key.GetValue('CodeBase') -ne 'file:///D:/Other/OfficeImageViewer.dll'){throw '旧注册被修改'}}finally{$key.Dispose()}
        Write-Output 'PASS conflicting installation rejected before any write'
    }finally{$base.DeleteSubKeyTree("Software\Classes\CLSID\$clsid",$false);$base.Dispose()}
}
Assert-Clean
try{
    & (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope
    & (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope
    foreach($view in $views){
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
        try{
            foreach($app in @('Word','PowerPoint')){
                $key=$base.OpenSubKey("Software\Microsoft\Office\$app\Addins\$progId")
                if(!$key){throw "$view $app registration missing"}
                try{if($key.GetValue('LoadBehavior') -ne 3){throw 'LoadBehavior invalid'}}finally{$key.Dispose()}
            }
        }finally{$base.Dispose()}
    }
    Write-Output 'PASS repeat install writes both host registrations in both views'
}finally{& (Join-Path $root 'uninstall.ps1') -Scope $Scope}
& (Join-Path $root 'uninstall.ps1') -Scope $Scope
Assert-Clean
Write-Output 'PASS repeat uninstall clears only product roots'
Write-Output 'Registry data tests: 4 passed; COM activation is NOT covered'

