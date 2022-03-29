Attribute VB_Name = "草图样条曲线"
Sub CATMain()
    
    '草图样条曲线绘制
    '***********************************
    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body

    Set oPart = CATIA.ActiveDocument.Part
    
   
    If Err.Number <> 0 Then

       
        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oBodies = oPart.Bodies
    Set oBody = oBodies.Item("零件几何体")
    
    '***********************************
    
    '定义草图平面
    Dim plnXY As Plane
    Set plnXY = oPart.OriginElements.PlaneYZ
    
    '再平面上添加草图
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plnXY)
    
    '添加草图工具
    Dim oFactory As Factory2D
    Set oFactory = oSketch.OpenEdition()
    
    '添加样条曲线控制点
    Dim p1 As ControlPoint2D, p2 As ControlPoint2D, p3 As ControlPoint2D, _
        p4 As ControlPoint2D, p5 As ControlPoint2D

    Set p1 = oFactory.CreateControlPoint(228.480423, 65.808388)
    Set p2 = oFactory.CreateControlPoint(172.887131, 35.341534)
    Set p3 = oFactory.CreateControlPoint(88.275314, 11.882068)
    Set p4 = oFactory.CreateControlPoint(14.965458, 57.887012)
    Set p5 = oFactory.CreateControlPoint(-45.820644, 27.420155)
    
    '按顺序连接控制点创建样条曲线
    Dim arrayOfObject1(4)
    Set arrayOfObject1(0) = p1
    Set arrayOfObject1(1) = p2
    Set arrayOfObject1(2) = p3
    Set arrayOfObject1(3) = p4
    Set arrayOfObject1(4) = p5
    Set oFactoryVariant = oFactory
    Set spline2D1 = oFactoryVariant.CreateSpline(arrayOfObject1)
    
    '退出草图
    oSketch.CloseEdition
    
    '更新
    oPart.Update

End Sub

