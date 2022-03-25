Attribute VB_Name = "fullname"
 '显示当前文件路径，未保存文件就显示名称
Sub CATMain()
    
    On Error Resume Next
    
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    '如果当前没有打开的文件
    If Err.Number <> 0 Then
    
        MsgBox "当前没有打开的文件"
        
    End If
    
    MsgBox oDoc.fullname
    
End Sub

