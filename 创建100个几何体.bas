Attribute VB_Name = "����100��������"
'Attribute VB_Name = "����100��������"
Sub CATMain()
    
        Dim oDoc As Document
        Dim oPart As Part
        Dim oBodies As Bodies
        Dim i As Integer
    
        Set oDoc = CATIA.ActiveDocument
        Set oPart = oDoc.Part
        Set oBodies = oPart.Bodies
        
    'ѭ��
        For i = 1 To 100
    
        Dim oBody
        Set oBody = oBodies.Add()
    
        oBody.Name = "������" & i
    
        Next
    
    MsgBox "�Ѿ��½���100����������"
    
    '����
        oPart.Update
    
End Sub


