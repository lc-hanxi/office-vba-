Attribute VB_Name = "模块1"
Sub Convert()
    '创建工作表
    Dim mySheetName As String, mySheetNameTest As String
    mySheetName = "Sheet2"
    
    On Error Resume Next
    mySheetNameTest = Worksheets(mySheetName).Name
    If Err.Number = 0 Then
        'MsgBox "The sheet named ''" & mySheetName & "'' DOES exist in this workbook."
    Else
        Err.Clear
        Worksheets.Add.Name = mySheetName
        'MsgBox "The sheet named ''" & mySheetName & "'' did not exist in this workbook but it has been created now."
    End If
    
    '设置表头
    Worksheets(mySheetName).Cells(1, 1) = "日期"
    Worksheets(mySheetName).Cells(1, 2) = "姓名"
    Worksheets(mySheetName).Cells(1, 3) = "项目名称"
    Worksheets(mySheetName).Cells(1, 4) = "工时"
    
    '遍历Sheet1, 设置Sheet2的内容
    oldsheetname = "Sheet1"
    s2Row = 2
    For Row = 2 To 100000
    
        If Worksheets(oldsheetname).Cells(Row, 1) = "合计" Then
            Exit For
        End If
        
        For Col = 3 To 1000
        
            If Worksheets(oldsheetname).Cells(1, Col) = "合计" Then
                Exit For
            End If
            
            If Worksheets(oldsheetname).Cells(Row, Col) > 0 Then
                Worksheets(mySheetName).Cells(s2Row, 1) = Worksheets(oldsheetname).Cells(Row, 1)
                Worksheets(mySheetName).Cells(s2Row, 2) = Worksheets(oldsheetname).Cells(Row, 2)
                Worksheets(mySheetName).Cells(s2Row, 3) = Worksheets(oldsheetname).Cells(1, Col)
                Worksheets(mySheetName).Cells(s2Row, 4) = Worksheets(oldsheetname).Cells(Row, Col)
                s2Row = s2Row + 1
            
            End If
        
        Next Col
        
    Next Row
    Worksheets(mySheetName).Range("a2:ao" & s2Row).Sort key1:=Worksheets(mySheetName).Range("b2"), order1:=xlAscending, Header:=xlNo
    
End Sub
