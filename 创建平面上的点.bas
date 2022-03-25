Attribute VB_Name = "����ƽ���ϵĵ�"
Sub CATMain()
    
    '��ʼ��������
    '***************************************************
    On Error Resume Next
    
    Dim oPart As Part
    Dim oHBodies As HybridBodies
    Dim oHBody As HybridBody
    Dim oHSF As HybridShapeFactory
    
    Set oPart = CATIA.ActiveDocument.Part
    
    If Err.Number <> 0 Then

        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oHBodies = oPart.HybridBodies
    Set oHSF = oPart.HybridShapeFactory
    Set oHBody = oHBodies.Add()
    '***************************************************
    
    '����ο�ƽ��
    Dim oPlane As Plane
    Set oPlane = oPart.OriginElements.PlaneYZ
    
    Dim oRefPlane As Reference
    Set oRefPlane = oPart.CreateReferenceFromObject(oPlane)
    
    '��ӵ�
    Dim oPoint5 As Point
    Set oPoint5 = oHSF.AddNewPointOnPlane(oRefPlane, 20, 60)
    
    oHBody.AppendHybridShape oPoint5
    
    '���¼���ͼ�μ�������Part�ĵ�
    oPart.UpdateObject oHBody
    oPart.Update
End Sub

