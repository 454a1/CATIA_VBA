Attribute VB_Name = "Module1"
'Attribute VB_Name = "originElements"
Sub CATMain()
    
    '��ȡ��ǰ�򿪵�����ĵ�

        Dim oDoc As Document
        Dim oPart As Part

        Set oDoc = CATIA.ActiveDocument
        Set oPart = oDoc.Part

    'ͨ��originElements���Է�������ĵ��Ĳο�ƽ��
            
        Dim plnXYZ As Plane
        Set plnXYZ = oPart.originElements.PlaneXY

    '��ʾ����
    
        MsgBox plnXYZ.Name
    
End Sub



