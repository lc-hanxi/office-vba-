Sub �������ͼ()

    Dim rectBase As Shape
    Dim rectSub(100) As Shape
    Dim connector(100) As Shape
    
    
    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150
    Dim content As Variant
    Dim clen As Long
    Dim basex, basey, cstart As Double
      
    content = Array("���ݻ㼯�ַ�����", "���ݽ���", "���ݴ���", "����ת��", "���ݷַ�")
    clen = UBound(content) '�����������±�

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("��ͼ��������ʽ")
    
    basex = 4
    basey = 8
    Set rectBase = ActiveWindow.Page.Drop(Application.Documents.Item("BASIC_M.VSSX").Masters.ItemU("Rectangle"), basex, basey)
    rectBase.CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "60 mm"
    rectBase.CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "15 mm"
    rectBase.CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = "249"
    rectBase.CellsSRC(visSectionCharacter, 0, visCharacterAsianFont).FormulaU = "249"
    rectBase.CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "14 pt"
    rectBase.Characters = content(0)
    
    cstart = basex - (clen - 1) / 2
    
    For i = 1 To clen
        
        '�����Ӿ��Σ����ô�С
        Set rectSub(i) = ActiveWindow.Page.Drop(Application.Documents.Item("BASIC_M.VSSX").Masters.ItemU("Rectangle"), cstart + i - 1, basey - 2)
        rectSub(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "15 mm"
        rectSub(i).CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "60 mm"
        
        '������������ݸ�ʽ
        rectSub(i).Characters = content(i)
        rectSub(i).CellsSRC(visSectionCharacter, 0, visCharacterFont).FormulaU = "249"
        rectSub(i).CellsSRC(visSectionCharacter, 0, visCharacterAsianFont).FormulaU = "249"
        rectSub(i).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "14 pt"
        
        '����
        Set connector(i) = ActiveWindow.Page.Drop(Application.ConnectorToolDataObject, 0, 0)
        Dim vsoCell1 As Visio.Cell
        Dim vsoCell2 As Visio.Cell
        Set vsoCell1 = connector(i).CellsU("BeginX")
        Set vsoCell2 = rectBase.CellsSRC(7, 0, 0)
        vsoCell1.GlueTo vsoCell2
        Set vsoCell1 = connector(i).CellsU("EndX")
        Set vsoCell2 = rectSub(i).CellsSRC(7, 2, 0)
        vsoCell1.GlueTo vsoCell2

        
    Next i
    
    Application.EndUndoScope UndoScopeID1, True
    
    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices
    
End Sub
