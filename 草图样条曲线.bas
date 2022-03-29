Attribute VB_Name = "��ͼ��������"
Sub CATMain()
    
    '��ͼ�������߻���
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
    Set oBody = oBodies.Item("���������")
    
    '***********************************
    
    '�����ͼƽ��
    Dim plnXY As Plane
    Set plnXY = oPart.OriginElements.PlaneYZ
    
    '��ƽ������Ӳ�ͼ
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plnXY)
    
    '��Ӳ�ͼ����
    Dim oFactory As Factory2D
    Set oFactory = oSketch.OpenEdition()
    
    '����������߿��Ƶ�
    Dim p1 As ControlPoint2D, p2 As ControlPoint2D, p3 As ControlPoint2D, _
        p4 As ControlPoint2D, p5 As ControlPoint2D

    Set p1 = oFactory.CreateControlPoint(228.480423, 65.808388)
    Set p2 = oFactory.CreateControlPoint(172.887131, 35.341534)
    Set p3 = oFactory.CreateControlPoint(88.275314, 11.882068)
    Set p4 = oFactory.CreateControlPoint(14.965458, 57.887012)
    Set p5 = oFactory.CreateControlPoint(-45.820644, 27.420155)
    
    '��˳�����ӿ��Ƶ㴴����������
    Dim arrayOfObject1(4)
    Set arrayOfObject1(0) = p1
    Set arrayOfObject1(1) = p2
    Set arrayOfObject1(2) = p3
    Set arrayOfObject1(3) = p4
    Set arrayOfObject1(4) = p5
    Set oFactoryVariant = oFactory
    Set spline2D1 = oFactoryVariant.CreateSpline(arrayOfObject1)
    
    '�˳���ͼ
    oSketch.CloseEdition
    
    '����
    oPart.Update

End Sub

