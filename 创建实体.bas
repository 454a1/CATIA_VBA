Attribute VB_Name = "创建实体"
Sub CATMain()
    
     On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Dim oSF As ShapeFactory
    Dim oPad As Pad
    

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
    
    '画圆
    Dim oCCircle As Circle2D
    Set oCCircle = oFactory2D.CreateClosedCircle(-30, -50, 15)
    
    oSketch.CloseEdition
    
    '设置Pad特征的高度为20mm
    Set oSF = oPart.ShapeFactory
    Set oPad = oSF.AddNewPad(oSketch, 20)
    
    oPart.Update
    
End Sub

