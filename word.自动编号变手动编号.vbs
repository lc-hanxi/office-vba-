Sub 自动编号变手动编号()

    Dim p As Paragraph
    Dim mr As Range
    Dim arr() As String
    Dim i As Integer
    
    ReDim arr(1 To ActiveDocument.Paragraphs.Count + 1)
    
    i = 1
    For Each p In ActiveDocument.Paragraphs
        
        If p.Range.ListFormat.ListValue <> 0 Then
            arr(i) = p.Range.ListFormat.ListString
            i = i + 1
        End If
        
    Next p
    
    i = 1
    For Each p In ActiveDocument.Paragraphs
        
        If p.Range.ListFormat.ListValue <> 0 Then
            p.Range.ListFormat.ApplyNumberDefault
            p.Range.InsertBefore arr(i) + " "
            i = i + 1        
        End If
        
    Next p
     
End Sub