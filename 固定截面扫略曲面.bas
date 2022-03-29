Attribute VB_Name = "�̶�����ɨ������"
Sub CATMain()
    
    '�̶�����ɨ�������
    '**************************************
    On Error Resume Next
    
    Dim oPart As Part
    Dim oHBodies As HybridBodies
    Dim oHBody As HybridBody
    Dim oHSF As HybridShapeFactory
    Dim oBodies As Bodies
    Dim oBody As Body
    
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
    Set oBodies = oPart.Bodies
    Set oBody = oBodies.Item("���������")

    '**************************************
    '�������棬��ԲΪ����
    Dim point1 As Point
    Set point1 = oHSF.AddNewPointCoord(0#, 0#, 0#)

    oHBody.AppendHybridShape point1

    Dim oRefPoint1 As Reference
    Set oRefPoint1 = oPart.CreateReferenceFromObject(point1)
    '���ص�
    oHSF.GSMVisibility oRefPoint1, 0
    
    'ȷ��Բ���ڵ�ƽ��
    Dim plane1 As Plane
    Set plane1 = oPart.OriginElements.PlaneZX

    Dim oRefPlane1 As Reference
    Set oRefPlane1 = oPart.CreateReferenceFromObject(plane1)
    'AddNewCircleCtrRad��Բ
    Dim oCircle As HybridShapeCircle
    Set oCircle = oHSF.AddNewCircleCtrRad(oRefPoint1, oRefPlane1, False, 20#)

    oHBody.AppendHybridShape oCircle
    '�������
    '**************************************
    '����������
    '�����ͼƽ��
    Dim plane2 As Plane
    Set plane2 = oPart.OriginElements.PlaneYZ
    
    '�ڲο�ƽ������Ӳ�ͼ
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plane2)

    Dim oFactory As Factory2D
    Set oFactory = oSketch.OpenEdition()
    
    '����������߿��Ƶ�
    Dim p1 As ControlPoint2D, p2 As ControlPoint2D, p3 As ControlPoint2D, _
        p4 As ControlPoint2D

    Set p1 = oFactory.CreateControlPoint(0, 0)
    Set p2 = oFactory.CreateControlPoint(48.42411, 2.698587)
    Set p3 = oFactory.CreateControlPoint(82.875099, -14.21253)
    Set p4 = oFactory.CreateControlPoint(143.47995, -11.334044)
    
    '��˳�����ӿ��Ƶ㴴����������
    Dim arrayOfObject1(3)
    Set arrayOfObject1(0) = p1
    Set arrayOfObject1(1) = p2
    Set arrayOfObject1(2) = p3
    Set arrayOfObject1(3) = p4
    Set oFactoryVariant = oFactory
    Set spline2D1 = oFactoryVariant.CreateSpline(arrayOfObject1)
    
    oSketch.CloseEdition
    '���������
    '**************************************
    Dim oRefCle As Reference
    Set oRefCle = oPart.CreateReferenceFromObject(oCircle)

    Dim oRefSketch As Reference
    Set oRefSketch = oPart.CreateReferenceFromObject(oSketch)
    
    'ɨ��
    Dim oSE As HybridShapeSweepExplicit
    Set oSE = oHSF.AddNewSweepExplicit(oRefCle, oRefSketch)
    
    '���ؽ����������
    oHSF.GSMVisibility oRefCle, 0
    oHSF.GSMVisibility oRefSketch, 0

    oHBody.AppendHybridShape oSE


    oPart.Update

End Sub

