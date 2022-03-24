Attribute VB_Name = "Module1"
'Attribute VB_Name = "originElements"
Sub CATMain()
    
    '获取当前打开的零件文档

        Dim oDoc As Document
        Dim oPart As Part

        Set oDoc = CATIA.ActiveDocument
        Set oPart = oDoc.Part

    '通过originElements属性访问零件文档的参考平面
            
        Dim plnXYZ As Plane
        Set plnXYZ = oPart.originElements.PlaneXY

    '显示名称
    
        MsgBox plnXYZ.Name
    
End Sub



