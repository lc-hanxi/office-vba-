Sub 选中段落自动编号()
    Dim pc As Integer
    Dim id As Integer
    pc = Selection.Paragraphs.Count
    Selection.Collapse Direction:=wdCollapseStart
    id = Asc("A")
    For i = 1 To pc
        Selection.TypeText "（" + Chr(id) + "）"
        Selection.Move wdParagraph, 1
        id = id + 1
    Next i
End Sub