[CmdletBinding()]
param([ValidateSet('CurrentUser','AllUsers')][string]$Scope='CurrentUser')
if ($Scope -eq 'CurrentUser' -and [Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw "请使用普通权限 PowerShell 验证每用户 COM 激活。" }
$ErrorActionPreference='Stop'
$hive=if($Scope -eq 'AllUsers'){[Microsoft.Win32.RegistryHive]::LocalMachine}else{[Microsoft.Win32.RegistryHive]::CurrentUser}
$root=Split-Path $PSScriptRoot -Parent
$views=@([Microsoft.Win32.RegistryView]::Registry32)
if([Environment]::Is64BitOperatingSystem){$views += [Microsoft.Win32.RegistryView]::Registry64}
foreach($view in $views){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try{
        foreach($path in @('Software\Classes\CLSID\{26C7344A-8129-4C5B-A462-034F9B5FC712}','Software\Classes\OfficeImageViewer.AddIn','Software\Microsoft\Office\Word\Addins\OfficeImageViewer.AddIn','Software\Microsoft\Office\PowerPoint\Addins\OfficeImageViewer.AddIn')){
            $key=$base.OpenSubKey($path)
            if($key){$key.Dispose();throw '检测到已有安装；测试拒绝覆盖，请先卸载或换干净用户环境。'}
        }
    }finally{$base.Dispose()}
}
try {
    & (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope
    & (Join-Path $root 'install.ps1') -AllowMissingOffice -Scope $Scope
    $shells=@((Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'))
    if([Environment]::Is64BitOperatingSystem){$shells += Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'}
    foreach($shell in $shells){
        & $shell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'ComActivation.ps1')
        if($LASTEXITCODE -ne 0){throw "COM 激活验证失败：$shell"}
    }
}finally {
    & (Join-Path $root 'uninstall.ps1') -Scope $Scope
}
foreach($view in $views){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try{
        $key=$base.OpenSubKey('Software\Classes\CLSID\{26C7344A-8129-4C5B-A462-034F9B5FC712}')
        if($key){$key.Dispose();throw '卸载后残留 COM 注册'}
    }finally{$base.Dispose()}
}
Write-Output 'PASS idempotent user registration and cleanup'



