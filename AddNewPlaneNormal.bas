Attribute VB_Name = "AddNewPlaneNormal"
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
    
    '创建样条曲线的点
    Dim oPoint1 As Point, oPoint2 As Point, oPoint3 As Point, _
        oPoint4 As Point
    
    Set oPoint1 = oHSF.AddNewPointCoord(0, 2, 3)
    Set oPoint2 = oHSF.AddNewPointCoord(10, 5, 8)
    Set oPoint3 = oHSF.AddNewPointCoord(8, 9, 10)
    Set oPoint4 = oHSF.AddNewPointCoord(5, 25, 10)
    oHBody.AppendHybridShape oPoint1
    oHBody.AppendHybridShape oPoint2
    oHBody.AppendHybridShape oPoint3
    oHBody.AppendHybridShape oPoint4
    
    '以创建的点为参考
    Dim oRefPoint1 As Reference, oRefPoint2 As Reference, oRefPoint3 As Reference, _
        oRefPoint4 As Reference
    Set oRefPoint1 = oPart.CreateReferenceFromObject(oPoint1)
    Set oRefPoint2 = oPart.CreateReferenceFromObject(oPoint2)
    Set oRefPoint3 = oPart.CreateReferenceFromObject(oPoint3)
    Set oRefPoint4 = oPart.CreateReferenceFromObject(oPoint4)
    '隐藏元素
    oHSF.GSMVisibility oRefPoint1, 0
    oHSF.GSMVisibility oRefPoint2, 0
    oHSF.GSMVisibility oRefPoint3, 0
    
    Dim oHBSpline As HybridShapeSpline
    Set oHBSpline = oHSF.AddNewSpline()
    
    '创建样条曲线
    oHBSpline.AddPointWithConstraintExplicit oRefPoint1, Nothing, -1#, 1, Nothing, 0#
    oHBSpline.AddPointWithConstraintExplicit oRefPoint2, Nothing, -1#, 1, Nothing, 0#
    oHBSpline.AddPointWithConstraintExplicit oRefPoint3, Nothing, -1#, 1, Nothing, 0#
    
    oHBody.AppendHybridShape oHBSpline
    
    Dim oRefCurve As Reference
    Set oRefCurve = oPart.CreateReferenceFromObject(oHBSpline)
    '创建曲线上的点
    Dim oPoint5 As Point
    Set oPoint5 = oHSF.AddNewPointOnCurveFromDistance(oRefCurve, 25, False)
    
    oHBody.AppendHybridShape oPoint5
    
    Dim oRefPoint5 As Reference
    Set oRefPoint5 = oPart.CreateReferenceFromObject(oPoint5)
    
    '创建通过点oPoint4，且与曲线oHBSpline垂直的平面
    Dim oPlane1 As Plane, oPlane2 As Plane
    Set oPlane1 = oHSF.AddNewPlaneNormal(oRefCurve, oRefPoint5)
    '通过的点不一定要在曲线上
    Set oPlane2 = oHSF.AddNewPlaneNormal(oRefCurve, oRefPoint4)
    oHBody.AppendHybridShape oPlane1
    oHBody.AppendHybridShape oPlane2
    
    '更新几何图形集，更新Part文档
    oPart.UpdateObject oHBody
    oPart.Update
End Sub

