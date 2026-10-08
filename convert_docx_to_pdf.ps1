param(
    [string]$DocxPath,
    [string]$PdfPath
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $fullDocx = (Resolve-Path $DocxPath).Path
    $doc = $word.Documents.Open($fullDocx)
    
    # Update all fields like TOC, page numbers
    foreach ($field in $doc.Fields) {
        $field.Update()
    }
    
    # Save as PDF (Format 17 is wdFormatPDF)
    $fullPdf = [System.IO.Path]::GetFullPath($PdfPath)
    $doc.SaveAs([ref]$fullPdf, [ref]17)
    $doc.Close([ref]0)
    Write-Output "Successfully converted $fullDocx to $fullPdf"
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
