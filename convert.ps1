
 = New-Object -ComObject Word.Application
.Visible = False
 = Join-Path (Get-Location).Path 'test_doc.docx'
 = Join-Path (Get-Location).Path 'test_doc.pdf'
 = .Documents.Open()
.SaveAs([ref], [ref]17)
.Close(0)
.Quit()
