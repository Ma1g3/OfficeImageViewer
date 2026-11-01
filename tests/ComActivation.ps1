$ErrorActionPreference='Stop'
try {
    $addin=New-Object -ComObject OfficeImageViewer.AddIn
    $xml=[xml]$addin.GetCustomUI('Microsoft.Word.Document')
    if($xml.DocumentElement.LocalName -ne 'customUI'){throw 'Ribbon XML invalid'}
    $pointer=[Runtime.InteropServices.Marshal]::GetIUnknownForObject($addin)
    try {
        foreach($id in @('B65AD801-ABAF-11D0-BB8B-00A0C90F2744','000C0396-0000-0000-C000-000000000046')){
            $guid=[Guid]$id
            $result=[IntPtr]::Zero
            $hr=[Runtime.InteropServices.Marshal]::QueryInterface($pointer,[ref]$guid,[ref]$result)
            if($hr -ne 0){throw "QueryInterface failed: $id ($hr)"}
            [void][Runtime.InteropServices.Marshal]::Release($result)
        }
    }finally{[void][Runtime.InteropServices.Marshal]::Release($pointer)}
    if([Runtime.InteropServices.Marshal]::IsComObject($addin)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)}
    Write-Output "PASS COM activation + Ribbon + interfaces ($([IntPtr]::Size*8)-bit)"
    exit 0
}catch{Write-Output $_;exit 1}

