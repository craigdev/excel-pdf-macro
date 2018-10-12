Sub PDFTabsBetweenStartAndEnd()

Dim X As Long
  
For X = Sheets("PDF - Start").Index + 1 To Sheets("PDF - End").Index - 1
If Sheets(X).Visible = True Then
    Sheets(X).Select False
End If
Next
  
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=ActiveWorkbook.Path & "\Insert Name Here.pdf", _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=False
        
Sheets("PDF - Start").Select

End Sub

