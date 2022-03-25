Attribute VB_Name = "两点创建直线"
Sub CATMain()
    
    '初始基本操作
    '***************************************************
    On Error Resume Next
    
    Dim oPart As Part
    Dim oHBodies As HybridBodies
    Dim oHBody As HybridBody
    Dim oHSF As HybridShapeFactory
    
    Set oPart = CATIA.ActiveDocument.Part
    
    If Err.Number <> 0 Then

        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oHBodies = oPart.HybridBodies
    Set oHSF = oPart.HybridShapeFactory
    Set oHBody = oHBodies.Add()
    '***************************************************
    
    '创建点
    Dim oPoint1 As Point, oPoint2 As Point
    
    Set oPoint1 = oHSF.AddNewPointCoord(0, 2, 3)
    Set oPoint2 = oHSF.AddNewPointCoord(10, 5, 8)
    oHBody.AppendHybridShape oPoint1
    oHBody.AppendHybridShape oPoint2
    
    Dim oRefPoint1 As Reference, oRefPoint2 As Reference
    Set oRefPoint1 = oPart.CreateReferenceFromObject(oPoint1)
    Set oRefPoint2 = oPart.CreateReferenceFromObject(oPoint2)
    oHSF.GSMVisibility oRefPoint1, 0
    oHSF.GSMVisibility oRefPoint2, 0
    
    '创建直线
    Dim oLine As Line
    Set oLine = oHSF.AddNewLinePtPt(oRefPoint1, oRefPoint2)
    oHBody.AppendHybridShape oLine
    
    '更新几何图形集，更新Part文档
    oPart.UpdateObject oHBody
    oPart.Update

End Sub

