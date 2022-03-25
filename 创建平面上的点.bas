Attribute VB_Name = "创建平面上的点"
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
    
    '定义参考平面
    Dim oPlane As Plane
    Set oPlane = oPart.OriginElements.PlaneYZ
    
    Dim oRefPlane As Reference
    Set oRefPlane = oPart.CreateReferenceFromObject(oPlane)
    
    '添加点
    Dim oPoint5 As Point
    Set oPoint5 = oHSF.AddNewPointOnPlane(oRefPlane, 20, 60)
    
    oHBody.AppendHybridShape oPoint5
    
    '更新几何图形集，更新Part文档
    oPart.UpdateObject oHBody
    oPart.Update
End Sub

