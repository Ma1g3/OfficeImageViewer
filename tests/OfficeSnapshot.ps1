[CmdletBinding()]
param([ValidateSet('Word','PowerPoint')][string]$HostApp,[Parameter(Mandatory=$true)][string]$ExpectedFile,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSEdition -eq 'Core'){throw 'Use Windows PowerShell 5.1.'}
Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public static class OfficeClipboardProbe { [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber(); }'
$app=[Runtime.InteropServices.Marshal]::GetActiveObject($HostApp+'.Application')
$doc=if($HostApp -eq 'Word'){$app.ActiveDocument}else{$app.ActivePresentation}
if(![string]::Equals($doc.FullName,[IO.Path]::GetFullPath($ExpectedFile),[StringComparison]::OrdinalIgnoreCase)){throw 'Active document is not the expected test fixture.'}
$shapes=@()
if($HostApp -eq 'Word'){
    foreach($shape in $doc.InlineShapes){$shapes += [ordered]@{Kind='Inline';Width=$shape.Width;Height=$shape.Height;Type=$shape.Type;CropLeft=$shape.PictureFormat.CropLeft;CropTop=$shape.PictureFormat.CropTop}}
    foreach($shape in $doc.Shapes){$shapes += [ordered]@{Kind='Floating';Id=$shape.ID;Name=$shape.Name;Width=$shape.Width;Height=$shape.Height;Left=$shape.Left;Top=$shape.Top;Rotation=$shape.Rotation}}
}else{
    foreach($slide in $doc.Slides){foreach($shape in $slide.Shapes){$shapes += [ordered]@{SlideId=$slide.SlideID;Id=$shape.Id;Type=$shape.Type;Width=$shape.Width;Height=$shape.Height;Left=$shape.Left;Top=$shape.Top;Rotation=$shape.Rotation;CropLeft=if($shape.Type -eq 13){$shape.PictureFormat.CropLeft}else{$null};CropTop=if($shape.Type -eq 13){$shape.PictureFormat.CropTop}else{$null};CropRight=if($shape.Type -eq 13){$shape.PictureFormat.CropRight}else{$null};CropBottom=if($shape.Type -eq 13){$shape.PictureFormat.CropBottom}else{$null}}}}
}
$stream=[IO.File]::Open($ExpectedFile,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
$sha=[Security.Cryptography.SHA256]::Create()
try{$fileHash=[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$stream.Dispose();$sha.Dispose()}
$state=[ordered]@{Host=$HostApp;Saved=[Convert]::ToBoolean($doc.Saved);FileSha256=$fileHash;ClipboardSequence=[OfficeClipboardProbe]::GetClipboardSequenceNumber();Shapes=$shapes;AddinConnected=[bool]$app.COMAddIns.Item('OfficeImageViewer.AddIn').Connect}
$state | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output ($HostApp+' saved='+$state.Saved+' clipboard='+$state.ClipboardSequence+' connected='+$state.AddinConnected)
