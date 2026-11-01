[CmdletBinding()]
param()
$ErrorActionPreference='Continue'
$framework=(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Release
Write-Output "Windows 64-bit: $([Environment]::Is64BitOperatingSystem)"
Write-Output ".NET Framework Release: $framework (需要 >= 528040)"
foreach($app in @('Word','PowerPoint')){
    Write-Output "$app installed: $(Test-Path "Registry::HKEY_CLASSES_ROOT\$app.Application")"
}
$views=@([Microsoft.Win32.RegistryView]::Registry32)
if([Environment]::Is64BitOperatingSystem){$views += [Microsoft.Win32.RegistryView]::Registry64}
foreach($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryHive]::LocalMachine)){
foreach($view in $views){
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
    try{
        $key=$base.OpenSubKey('Software\Classes\CLSID\{26C7344A-8129-4C5B-A462-034F9B5FC712}\InprocServer32')
        if($key){
            $uri=[Uri]$key.GetValue('CodeBase')
            Write-Output "$hive $view COM path: $($uri.LocalPath), exists: $(Test-Path -LiteralPath $uri.LocalPath)"
            $key.Dispose()
        }else{Write-Output "$hive $view COM registration: absent"}
        foreach($app in @('Word','PowerPoint')){
            $key=$base.OpenSubKey("Software\Microsoft\Office\$app\Addins\OfficeImageViewer.AddIn")
            if($key){Write-Output "$hive $view $app LoadBehavior: $($key.GetValue('LoadBehavior'))";$key.Dispose()}
            else{Write-Output "$hive $view $app add-in: absent"}
        }
    }finally{$base.Dispose()}
}
}

