[CmdletBinding(SupportsShouldProcess=$true)]
param([ValidateSet('CurrentUser','AllUsers')][string]$Scope='CurrentUser')
$ErrorActionPreference='Stop'
$hive=if($Scope -eq 'AllUsers'){[Microsoft.Win32.RegistryHive]::LocalMachine}else{[Microsoft.Win32.RegistryHive]::CurrentUser}
$scopeLabel=if($Scope -eq 'AllUsers'){'所有用户'}else{'当前用户'}
if($Scope -eq 'AllUsers' -and !$WhatIfPreference -and ![Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'AllUsers 安装/卸载需要管理员权限 PowerShell。'}
if(Get-Process WINWORD,POWERPNT -ErrorAction SilentlyContinue){throw '请先关闭 Word 和 PowerPoint。'}
$clsid='{26C7344A-8129-4C5B-A462-034F9B5FC712}'
$progId='OfficeImageViewer.AddIn'
if(!$PSCmdlet.ShouldProcess($progId,('移除'+$scopeLabel+'插件注册'))){return}
$views=@([Microsoft.Win32.RegistryView]::Registry32)
if([Environment]::Is64BitOperatingSystem){$views += [Microsoft.Win32.RegistryView]::Registry64}
foreach($view in $views){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try{
        foreach($app in @('Word','PowerPoint')){$base.DeleteSubKeyTree("Software\Microsoft\Office\$app\Addins\$progId",$false)}
        $base.DeleteSubKeyTree("Software\Classes\$progId",$false)
        $base.DeleteSubKeyTree("Software\Classes\CLSID\$clsid",$false)
    }finally{$base.Dispose()}
}
Write-Host ('已移除'+$scopeLabel+'的 OfficeImageViewer 注册。源码和发布包保留。')
