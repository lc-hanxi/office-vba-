Attribute VB_Name = "模块1"
Sub Sleep(T As Single)  ' T 参数的单位是 秒级
    Dim time1 As Single
    time1 = Timer
    Do
        DoEvents '转让控制权，以便让操作系统处理其它的事件
    Loop While Timer - time1 < T  ' T 参数的单位是 秒级
End Sub

'移动单元格
Sub CellMoveTo(rs As Integer, cs As Integer, re As Integer, ce As Integer)
    
    Worksheets("Sheet2").Cells(rs, cs).Select
    Selection.Cut
    
    Worksheets("Sheet2").Cells(re, ce).Select
    ActiveSheet.Paste

End Sub


'同一行两个单元格交换
Sub Swap(row As Integer, col1 As Integer, col2 As Integer)
    
    Call CellMoveTo(row, col1, row - 2, col1)
    Call Sleep(1)
    
    Call CellMoveTo(row, col2, row - 1, col2)
    Call Sleep(1)
    
    Dim i%, j%
    i = col1
    j = col2
    
    Do While i < col2
        
        Call CellMoveTo(row - 2, i, row - 2, i + 1)
        i = i + 1
        
        Call CellMoveTo(row - 1, j, row - 1, j - 1)
        j = j - 1
        
        Call Sleep(1)
    Loop
    
    Call CellMoveTo(row - 1, col1, row, col1)
    Call Sleep(1)
    
    Call CellMoveTo(row - 2, col2, row, col2)
    Call Sleep(1)
    
End Sub



Sub Color(row As Integer, col As Integer, clr As Long)
    
    Worksheets("Sheet2").Cells(row, col).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = clr
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub



'冒泡排序
Sub BubbleSort()

    Dim i%, j%, mend%, row%
    Dim clr1 As Long, clr2 As Long
    
    
    mend = 14
    row = 7
    clr1 = 5287936
    clr2 = 49407
    
    For i = 5 To 13
        For j = 5 To mend - 1
            Call Color(row, j, clr2)
            Call Color(row, j + 1, clr2)
            Call Sleep(1)
            
            If Worksheets("Sheet2").Cells(row, j).Value > Worksheets("Sheet2").Cells(row, j + 1).Value Then
                Call Swap(row, j, j + 1)
            End If
            
            Call Color(row, j, clr1)
            Call Color(row, j + 1, clr1)
            Call Sleep(1)
        Next j
        mend = mend - 1
    Next i

End Sub

