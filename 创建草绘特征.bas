Attribute VB_Name = "创建草绘特征"
Sub CATMain()
    
    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Dim dPi As Double
    
    dPi = 3.14159265358979

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
    
    '利用AbsoluteAxis获取轴系统
    'Dim oHRef As Line2D, oVRef As Line2D
    
    'Set oHRef = oSketch.AbsoluteAxis.HorizontalReference
    'Set oVRef = oSketch.AbsoluteAxis.VerticalReference
    
    '获取Factory2D，利用Factory2D可以创建草图元素（画圆，画直线等）
    Dim oFactory2D As Factory2D
    Set oFactory2D = oSketch.OpenEdition
    
    '创建点
    Dim oPoint As Point2D
    Set oPoint = oFactory2D.CreatePoint(10, 20)
    
    '创建直线
    Dim oLine As Line2D
    Set oLine = oFactory2D.CreateLine(0, 0, 30, 65)
    
    '创建圆弧
    Dim oCircle As Circle2D
    Set oCircle = oFactory2D.CreateCircle(20, 30, 10, dPi / 4, dPi / 2)
    
    '创建整圆
    Dim oCCircle As Circle2D
    Set oCCircle = oFactory2D.CreateClosedCircle(-30, -50, 15)
    
    oPart.Update
    
End Sub

