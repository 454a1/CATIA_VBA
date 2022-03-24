Attribute VB_Name = "���Լ��"
Sub CATMain()

    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body

    '���õ�ǰ�ĵ�
    Set oPart = CATIA.ActiveDocument.Part
    
    '�����ǰû���ĵ���
    If Err.Number <> 0 Then

       '�½��ĵ�
        Dim oDoc As Document
        Set oDoc = CATIA.Documents.Add("Part")
        Set oPart = oDoc.Part

    End If

    On Error GoTo 0
    
    Set oBodies = oPart.Bodies
    Set oBody = oBodies.Item("���������")
      
    'ѡȡ�ο�ƽ��
    Dim plnXY As Plane
    Set plnXY = oPart.originElements.PlaneYZ
    
    '�ڲο�ƽ������Ӳ�ͼ
    Dim oSketch As Sketch
    Set oSketch = oBody.Sketches.Add(plnXY)
    
    '��ȡFactory2D������Factory2D���Դ�����ͼԪ�أ���Բ����ֱ�ߵȣ�
    Dim oFactory2D As Factory2D
    Set oFactory2D = oSketch.OpenEdition
    
    
    '����ֱ��
    Dim oLine As Line2D
    Set oLine = oFactory2D.CreateLine(0, 0, 30, 65)
    
    Dim oLineH As Line2D
    Set oLineH = oSketch.AbsoluteAxis.HorizontalReference
    
    '�����ο�
    Dim oRef1 As Reference, oRef2 As Reference
    
    Set oRef1 = oPart.CreateReferenceFromObject(oLine)
    Set oRef2 = oPart.CreateReferenceFromObject(oLineH)
    
    '����Լ��
    Dim oConstraints As Constraints, oConstraint As Constraint
    Set oConstraints = oSketch.Constraints
    
    Set oConstraint = oConstraints.AddBiEltCst(catCstTypeAngle, oRef1, oRef2)
    oConstraint.Dimension.Value = 30
    
    oPart.Update
    
End Sub

