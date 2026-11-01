[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core' -or $PSVersionTable.PSVersion.Major -ne 5){throw 'Run this check with Windows PowerShell 5.1.'}
$root=Split-Path $PSScriptRoot -Parent
$files=@(Get-ChildItem -LiteralPath $root -Filter '*.ps1' -Recurse | Where-Object {
    $relative=$_.FullName.Substring($root.Length).TrimStart([IO.Path]::DirectorySeparatorChar)
    $relative -notmatch '^(\.git|\.local|bin|obj|dist|work|validation-results)[\\/]'
})
foreach($file in $files){
    $bytes=[IO.File]::ReadAllBytes($file.FullName)
    if($bytes.Length -lt 3 -or $bytes[0] -ne 239 -or $bytes[1] -ne 187 -or $bytes[2] -ne 191){throw "Missing UTF-8 BOM: $($file.Name)"}
    $tokens=$null;$issues=$null
    [void][Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$issues)
    if($issues.Count){throw "PS5.1 parse failure: $($file.Name) $issues"}
    Write-Output "PASS PS5.1 encoding and parsing: $($file.Name)"
}
Write-Output "Script compatibility: $($files.Count) files passed"
