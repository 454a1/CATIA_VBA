Attribute VB_Name = "添加约束"
Sub CATMain()

    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body

    '利用当前文档
    Set oPart = CATIA.ActiveDocument.Part
    
    '如果当前没有文档打开
    If Err.Number <> 0 Then

       '新建文档
        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oBodies = oPart.Bodies
    Set oBody = oBodies.Item("零件几何体")
      
    '选取参考平面
    Dim plnXY As Plane
    Set plnXY = oPart.originElements.PlaneYZ
    
    '在参考平面上添加草图
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plnXY)
    
    '获取Factory2D，利用Factory2D可以创建草图元素（画圆，画直线等）
    Dim oFactory2D As Factory2D
    Set oFactory2D = oSketch.OpenEdition
    
    
    '创建直线
    Dim oLine As Line2D
    Set oLine = oFactory2D.CreateLine(0, 0, 30, 65)
    
    Dim oLineH As Line2D
    Set oLineH = oSketch.AbsoluteAxis.HorizontalReference
    
    '建立参考
    Dim oRef1 As Reference, oRef2 As Reference
    
    Set oRef1 = oPart.CreateReferenceFromObject(oLine)
    Set oRef2 = oPart.CreateReferenceFromObject(oLineH)
    
    '创建约束
    Dim oConstraints As Constraints, oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeAngle, oRef1, oRef2)
    oConstraint.Dimension.Value = 30
    
    oPart.Update
    
End Sub

