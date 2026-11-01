[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$DllPath,
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\work\ppt-saved-state')
)
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -ne 5) { throw 'Requires Windows PowerShell 5.1.' }
$dll = (Resolve-Path -LiteralPath $DllPath).Path
$output = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) { throw 'Use a new output directory; existing files are never overwritten.' }
New-Item -ItemType Directory -Path $output | Out-Null
Add-Type -AssemblyName System.Drawing
[Reflection.Assembly]::LoadFrom($dll) | Out-Null
$image = Join-Path $output 'fixture.png'
$bitmap = New-Object Drawing.Bitmap 320,200
$graphics = [Drawing.Graphics]::FromImage($bitmap)
try { $graphics.Clear([Drawing.Color]::CornflowerBlue); $bitmap.Save($image, [Drawing.Imaging.ImageFormat]::Png) }
finally { $graphics.Dispose(); $bitmap.Dispose() }
function FileHash([string]$path) {
    $stream = [IO.File]::Open($path, 'Open', 'Read', [IO.FileShare]::ReadWrite)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
    finally { $stream.Dispose(); $sha.Dispose() }
}
function Assert($condition, [string]$message) { if (!$condition) { throw $message } }
$app = $null; $doc = $null; $adapter = $null; $ownsSession = $false
$results = @()
try {
    $app = New-Object -ComObject PowerPoint.Application
    # PowerPoint may reuse an existing process. Never touch an existing presentation.
    Assert ($app.Presentations.Count -eq 0) 'Existing presentations found; refusing to test.'
    $ownsSession = $true
    $app.Visible = -1
    $adapter = New-Object OfficeImageViewer.PowerPointHostAdapter -ArgumentList $app
    foreach ($format in @(24,25)) {
        foreach ($fontSetting in @(0,-1)) {
            $extension = if ($format -eq 25) { 'pptm' } else { 'pptx' }
            $path = Join-Path $output ("fixture-$format-$fontSetting.$extension")
            $doc = $app.Presentations.Add(-1)
            $slide = $doc.Slides.Add(1,12)
            $shape = $slide.Shapes.AddPicture($image,0,-1,100,100,320,200)
            $doc.SaveAs($path,$format,$fontSetting)
            $shape.Select(-1)
            Assert ([int]$doc.Saved -eq -1) 'Fixture must start saved.'
            $hash = FileHash $path
            $geometry = @($shape.Left,$shape.Top,$shape.Width,$shape.Height,$shape.PictureFormat.CropLeft,$shape.PictureFormat.CropRight) -join ','
            $tempBefore = @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'OfficeImageViewer-*' | ForEach-Object FullName)
            # Control: reproduce the old argument on the embedded-font fixture.
            $control = Join-Path $output "control-$format-$fontSetting.$extension"
            $doc.SaveCopyAs($control,$format,0)
            $oldSaved = [int]$doc.Saved
            # Only save our own fixture to reset the control side effect.
            $doc.Close(); $doc = $app.Presentations.Open($path,0,0,-1)
            $shape = $doc.Slides.Item(1).Shapes.Item(1)
            $shape.Select(-1)
            Assert ([int]$doc.Saved -eq -1) 'Reopened fixture must start saved.'
            for ($i=0; $i -lt 3; $i++) {
                $payload = $adapter.TryGetSelectedImage()
                Assert ($payload.Bytes.Length -gt 0) 'Preview bytes missing.'
                Assert ([int]$doc.Saved -eq -1) 'Preview dirtied a saved fixture.'
            }
            Assert ((FileHash $path) -eq $hash) 'Original file hash changed.'
            Assert ((@($shape.Left,$shape.Top,$shape.Width,$shape.Height,$shape.PictureFormat.CropLeft,$shape.PictureFormat.CropRight) -join ',') -eq $geometry) 'Image geometry changed.'
            # Real content edit, not an artificial Saved assignment.
            $null = $doc.Slides.Item(1).Shapes.AddShape(1,10,10,20,20)
            Assert ([int]$doc.Saved -eq 0) 'Fixture edit should be unsaved.'
            $shape.Select(-1)
            $null = $adapter.TryGetSelectedImage()
            Assert ([int]$doc.Saved -eq 0) 'Preview cleared real unsaved edits.'
            Assert ($doc.Slides.Item(1).Shapes.Count -eq 2) 'Unsaved content lost.'
            $tempAfter = @(Get-ChildItem -LiteralPath ([IO.Path]::GetTempPath()) -Directory -Filter 'OfficeImageViewer-*' | ForEach-Object FullName)
            Assert (@($tempAfter | Where-Object { $tempBefore -notcontains $_ }).Count -eq 0) 'Temporary preview directories remain.'
            $results += [ordered]@{Format=$format;EmbeddedFontsArgument=$fontSetting;OldArgumentSaved=$oldSaved;FixedSaved=-1;AfterRealEdit=0;OriginalSha256=$hash;RepeatedReads=3;TemporaryDirectoriesClean=$true}
            $doc.Close(); $doc = $null
        }
    }
    [ordered]@{OfficeVersion=[string]$app.Version;DllSha256=(FileHash $dll);Results=$results;Scope='Real COM adapter reads only; no mouse double-click, viewer or UI close-prompt acceptance'} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $output 'results.json') -Encoding UTF8
    $results | ConvertTo-Json -Depth 4
}
finally {
    if ($adapter) { $adapter.Dispose() }
    if ($doc) { $doc.Close() }
    if ($app -and $ownsSession -and $app.Presentations.Count -eq 0) { $app.Quit() }
    if ($app) { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null }
}
