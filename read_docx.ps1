Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipPath = "c:\Users\USER\Downloads\TripZoneTravels_Proposal.docx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
$entry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" }
if ($entry -eq $null) {
    Write-Host "word/document.xml not found"
} else {
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream)
    $xmlString = $reader.ReadToEnd()
    $reader.Close()
    $xml = [xml]$xmlString
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $nodes = $xml.SelectNodes("//w:p", $ns)
    $text = ""
    foreach ($node in $nodes) {
        $text += $node.InnerText + "`n"
    }
    Write-Host $text
}
$zip.Dispose()
