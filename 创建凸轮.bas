'Attribute VB_Name = "创建凸轮"
Sub CATMain()

    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Const Pi = 3.14159265358979

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
    
    Dim oLineH As Line2D, oPtO As Point2D
    Set oLineH = oSketch.AbsoluteAxis.HorizontalReference
    Set oPtO = oSketch.AbsoluteAxis.Origin
        
    Dim oLineConst As Line2D
    Set oLineConst = oFactory2D.CreateLine(0, 0, 50, 0)
    oLineConst.StartPoint = oPtO
    oLineConst.Construction = True
    
    Dim oCircle1 As Circle2D, oCircle2 As Circle2D
    Set oCircle1 = oFactory2D.CreateCircle(0, 0, 30, Pi / 2, -Pi / 2)
    oCircle1.CenterPoint = oLineConst.StartPoint
        
    Set oCircle2 = oFactory2D.CreateCircle(50, 0, 15, -Pi / 2, Pi / 2)
    oCircle2.CenterPoint = oLineConst.EndPoint
        
    Dim oL1 As Line2D, oL2 As Line2D
    Set oL1 = oFactory2D.CreateLine(0, 30, 50, 15)
    Set oL2 = oFactory2D.CreateLine(0, -30, 50, -15)
        
    oL1.StartPoint = oCircle1.StartPoint
    oL2.StartPoint = oCircle1.EndPoint
    oL1.EndPoint = oCircle2.EndPoint
    oL2.EndPoint = oCircle2.StartPoint
    
    '创建约束
    Dim oConstraints As Constraints, oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
        
    Dim oRefC1 As Reference, oRefC2 As Reference
    Set oRefC1 = oPart.CreateReferenceFromObject(oCircle1)
    Set oRefC2 = oPart.CreateReferenceFromObject(oCircle2)
        
    Dim oRefL1 As Reference, oRefL2 As Reference
    Set oRefL1 = oPart.CreateReferenceFromObject(oL1)
    Set oRefL2 = oPart.CreateReferenceFromObject(oL2)
        
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeTangency, oRefL1, oRefC1)
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeTangency, oRefL1, oRefC2)
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeTangency, oRefL2, oRefC1)
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeTangency, oRefL2, oRefC2)
    
    Set oConstraint = _
        oConstraints.AddMonoEltCst(catCstTypeRadius, oRefC1)
        oConstraint.Dimension.Value = 30
    Set oConstraint = _
        oConstraints.AddMonoEltCst(catCstTypeRadius, oRefC2)
        oConstraint.Dimension.Value = 15
        
     
        Dim oRefLH As Reference, oRefLC As Reference
        Set oRefLC = oPart.CreateReferenceFromObject(oLineConst)
        Set oRefLH = oPart.CreateReferenceFromObject(oLineH)
        
        Set oConstraint = _
            oConstraints.AddMonoEltCst(catCstTypeLength, oRefLC)
            oConstraint.Dimension.Value = 50
            
        Set oConstraint = _
            oConstraints.AddBiEltCst(catCstTypeAngle, oRefLH, oRefLC)
            oConstraint.Dimension.Value = 90 * -1
            
        oSketch.CloseEdition
        
        Dim oSF As ShapeFactory
        Set oSF = oPart.ShapeFactory
        Dim oPadCam As Pad
        Set oPadCam = oSF.AddNewPad(oSketch, 20)
        
        oPadCam.FirstLimit.Dimension.Value = 10 + 20
        oPadCam.SecondLimit.Dimension.Value = 10 * -1
        
    oPart.Update
    
        
End Sub

