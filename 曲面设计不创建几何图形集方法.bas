Attribute VB_Name = "曲面设计不创建几何图形集方法"
Sub CATMain()

On Error Resume Next
    
    Dim oPart As Part
    Dim oHSF As HybridShapeFactory
    Dim oBodies As bodies
    Dim oBody As Body
    
    Set oPart = CATIA.ActiveDocument.Part
    
    If Err.Number <> 0 Then

        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oHSF = oPart.HybridShapeFactory
    Set oBodies = oPart.bodies
    Set oBody = oBodies.Item("零件几何体")
    
    Dim Point1 As HybridShapePointCoord
    Set Point1 = oHSF.AddNewPointCoord(0#, 2#, 3#)
    
    '使用InsertHybridShape可以不添加几何图形集
    oBody.InsertHybridShape Point1


    Dim oRefPt As Reference
    Set oRefPt = oPart.CreateReferenceFromObject(Point1)
    oHSF.GSMVisibility oRefPt, 0
    
    '定义方向
    Dim dir As HybridShapeDirection
    Set dir = oHSF.AddNewDirectionByCoord(1#, 2#, 3#)

    Dim oLine As Line
    Set oLine = oHSF.AddNewLinePtDir(oRefPt, dir, 0#, 100#, False)

    oBody.InsertHybridShape oLine
    
    '更新
    oPart.Update

End Sub

