Attribute VB_Name = "fullname"
 '��ʾ��ǰ�ļ�·����δ�����ļ�����ʾ����
Sub CATMain()
    
    On Error Resume Next
    
    Dim oDoc As Document
    Set oDoc = CATIA.ActiveDocument
    
    '�����ǰû�д򿪵��ļ�
    If Err.Number <> 0 Then
    
        MsgBox "��ǰû�д򿪵��ļ�"
        
    End If
    
    MsgBox oDoc.fullname
    
End Sub

