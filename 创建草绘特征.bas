Attribute VB_Name = "�����ݻ�����"
Sub CATMain()
    
    On Error Resume Next
    
    Dim oPart As Part
    Dim oBodies As Bodies
    Dim oBody As Body
    Dim dPi As Double
    
    dPi = 3.14159265358979

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
    
    '����AbsoluteAxis��ȡ��ϵͳ
    'Dim oHRef As Line2D, oVRef As Line2D
    
    'Set oHRef = oSketch.AbsoluteAxis.HorizontalReference
    'Set oVRef = oSketch.AbsoluteAxis.VerticalReference
    
    '��ȡFactory2D������Factory2D���Դ�����ͼԪ�أ���Բ����ֱ�ߵȣ�
    Dim oFactory2D As Factory2D
    Set oFactory2D = oSketch.OpenEdition
    
    '������
    Dim oPoint As Point2D
    Set oPoint = oFactory2D.CreatePoint(10, 20)
    
    '����ֱ��
    Dim oLine As Line2D
    Set oLine = oFactory2D.CreateLine(0, 0, 30, 65)
    
    '����Բ��
    Dim oCircle As Circle2D
    Set oCircle = oFactory2D.CreateCircle(20, 30, 10, dPi / 4, dPi / 2)
    
    '������Բ
    Dim oCCircle As Circle2D
    Set oCCircle = oFactory2D.CreateClosedCircle(-30, -50, 15)
    
    oPart.Update
    
End Sub

