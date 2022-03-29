Attribute VB_Name = "拉伸曲线成曲面"
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
    
    '使用InsertHybridShape可以不添加几何图形集
    '创建样条曲线的点
    Dim oPoint1 As Point, oPoint2 As Point, oPoint3 As Point
    
    Set oPoint1 = oHSF.AddNewPointCoord(0, 2, 3)
    Set oPoint2 = oHSF.AddNewPointCoord(10, 5, 8)
    Set oPoint3 = oHSF.AddNewPointCoord(8, 9, 10)
    oBody.InsertHybridShape oPoint1
    oBody.InsertHybridShape oPoint2
    oBody.InsertHybridShape oPoint3
    
    '以创建的点为参考
    Dim oRefPoint1 As Reference, oRefPoint2 As Reference, oRefPoint3 As Reference
    Set oRefPoint1 = oPart.CreateReferenceFromObject(oPoint1)
    Set oRefPoint2 = oPart.CreateReferenceFromObject(oPoint2)
    Set oRefPoint3 = oPart.CreateReferenceFromObject(oPoint3)
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
    
    oBody.InsertHybridShape oHBSpline
    
    '将曲线设为参考
    Dim oRefCurve As Reference
    Set oRefCurve = oPart.CreateReferenceFromObject(oHBSpline)
    
    '定义方向
    Dim dir As HybridShapeDirection
    Set dir = oHSF.AddNewDirectionByCoord(0#, 0#, 1#)
    
    '拉伸曲线
    Dim oExtrude As HybridShapeExtrude
    Set oExtrude = oHSF.AddNewExtrude(oRefCurve, 20, 0, dir)
    
    oBody.InsertHybridShape oExtrude
    
    oPart.Update

End Sub

