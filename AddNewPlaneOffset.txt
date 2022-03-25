Attribute VB_Name = "AddNewPlaneOffset"
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
    '创建坐标点
    Dim oPoint1 As Point, oPoint2 As Point, oPoint3 As Point
    
    Set oPoint1 = oHSF.AddNewPointCoord(0, 2, 3)
    Set oPoint2 = oHSF.AddNewPointCoord(10, 5, 8)
    Set oPoint3 = oHSF.AddNewPointCoord(8, 9, 10)
    oHBody.AppendHybridShape oPoint1
    oHBody.AppendHybridShape oPoint2
    oHBody.AppendHybridShape oPoint3
    
    '是否隐藏创建的点？
    '以创建的点为参考
    Dim oRefPoint1 As Reference, oRefPoint2 As Reference, oRefPoint3 As Reference
    Set oRefPoint1 = oPart.CreateReferenceFromObject(oPoint1)
    Set oRefPoint2 = oPart.CreateReferenceFromObject(oPoint2)
    Set oRefPoint3 = oPart.CreateReferenceFromObject(oPoint3)
    '隐藏点
    oHSF.GSMVisibility oRefPoint1, 0
    oHSF.GSMVisibility oRefPoint2, 0
    oHSF.GSMVisibility oRefPoint3, 0
    
    '创建通过三点的平面
    Dim oPlane1 As Plane
    Set oPlane1 = oHSF.AddNewPlane3Points(oRefPoint1, oRefPoint2, oRefPoint3)
    oHBody.AppendHybridShape oPlane1
    
    Dim oRefPlane1 As Reference
    Set oRefPlane1 = oPart.CreateReferenceFromObject(oPlane1)
    
    '创建偏移平面
    Dim oPlane2 As Plane
    Set oPlane2 = oHSF.AddNewPlaneOffset(oRefPlane1, 200, False)
    oHBody.AppendHybridShape oPlane2
    
    oPart.Update
    
End Sub


