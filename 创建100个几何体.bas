Attribute VB_Name = "创建100个几何体"
'Attribute VB_Name = "创建100个几何体"
Sub CATMain()
    
        Dim oDoc As Document
        Dim oPart As Part
        Dim oBodies As Bodies
        Dim i As Integer
    
        Set oDoc = CATIA.ActiveDocument
        Set oPart = oDoc.Part
        Set oBodies = oPart.Bodies
        
    '循环
        For i = 1 To 100
    
        Dim oBody
        Set oBody = oBodies.Add()
    
        oBody.Name = "几何体" & i
    
        Next
    
    MsgBox "已经新建了100个几何体了"
    
    '更新
        oPart.Update
    
End Sub


